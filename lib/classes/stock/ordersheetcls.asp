<%
'####################################################
' Description :  주문서관리 클래스
' History : 이상구 생성
'####################################################

Class CBaljusumItem
	public FMasteridx
	public FDetailidx
	'public FBaljuid
	public FMakerid
	public FItemGubun
	public FItemId
	public FItemoption
	public FJupsuCount  '' 주문접수
	public FCount       '' 주문확인
	public FItemName
	public FItemOptionname
    public Frealstock   '' 실사재고(창고)
    public Fmwdiv
    public Fpreorderno
    public Fpreordernofix
    public FreipgoMayDate
    public Fshopid
    public Fshopname
    public Fbaljucode
	Public Fbaljucodecnt
    public Fupcheorderlinkcode

	public Fsmallimage
	public Foffimgsmall

	public Fdeliverytype
	public Fcentermwdiv

	Public FpriceCnt
	public FpreUnderCnt
	Public Fonbaljuitemno
	public Ftotbaljuitemno

    public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = Fsmallimage
		else
			GetImageSmall = Foffimgsmall
		end if
	end function

	public function GetDeliverTypeString()
		if Fmwdiv="U" then
			GetDeliverTypeString = "업배"
		else
			GetDeliverTypeString = ""
		end if
	end function

	public function GetMWDivString()
		if Fmwdiv="M" then
			GetMWDivString = "매입"
		elseif Fmwdiv="W" then
			GetMWDivString = "특정"
		else
			if (FItemGubun <> "10") then
				GetMWDivString = "오프"
			end if
		end if
	end function

	public function GetCenterMWDivString()
		if Fcentermwdiv="M" then
			GetCenterMWDivString = "매입"
		elseif Fcentermwdiv="W" then
			GetCenterMWDivString = "특정"
		else
			GetCenterMWDivString = "-"
		end if
	end function

	public function IsOnLineItem()
		IsOnLineItem = (Fitemgubun="10")
	end function

	public function IsUpchebeasong()
		if (Fdeliverytype="2") or (Fdeliverytype="9") or (Fdeliverytype="7") then
			IsUpchebeasong = true
		else
			IsUpchebeasong = false
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COrderSheetCheckItem
	public FIdx
	public Fbaljuid
	public Fbaljucode
	public FMakerid
	public Ftotalsellcash
	public Ftotalbuycash
	public Fcount
	public Ftotalsellcash2
	public Ftotalbuycash2
	public Fcount2
	public FRegDate
	public FScheduleDate
	public Fbaljuname
	public Fstatecd
	public Fipgodate
	public Fdivcode

	public function GetStateName()
		if Fstatecd=" " and fforeign_statecd="0" then
			GetStateName = "업체접수(견적요청)"
		elseif Fstatecd=" " and fforeign_statecd="3" then
			GetStateName = "업체접수확인"
		elseif Fstatecd=" " and fforeign_statecd="7" then
			GetStateName = "업체컨펌완료"
		elseif Fstatecd=" " and (isnull(fforeign_statecd) or fforeign_statecd="") then
			GetStateName = "주문서작성중"
		elseif Fstatecd="0" then
			GetStateName = "주문접수"
		elseif Fstatecd="1" then
			GetStateName = "주문확인"
		elseif Fstatecd="2" then
			GetStateName = "입금대기"
		elseif Fstatecd="5" then
			GetStateName = "배송준비"
		elseif Fstatecd="6" then
			GetStateName = "출고대기"
		elseif Fstatecd="7" then
			GetStateName = "출고완료"
		elseif Fstatecd="8" then
			GetStateName = "검품완료"
		elseif Fstatecd="9" then
			GetStateName = "입고완료"
		end if
	end function

	public function GetStateColor()
		if Fstatecd="0" then
			GetStateColor = "#00000"
		elseif Fstatecd="1" then
			GetStateColor = "#00AA00"
		elseif Fstatecd="2" then
			GetStateColor = "#0000AA"
		elseif Fstatecd="5" then
			GetStateColor = "#AAAA00"
		elseif Fstatecd="6" then
			GetStateColor = "#AA00AA"
		elseif Fstatecd="7" then
			GetStateColor = "#AA0000"
		elseif Fstatecd="8" then
			GetStateColor = "#33AAAA"
		elseif Fstatecd="9" then
			GetStateColor = "#AA33AA"
		elseif Fstatecd=" " then
			GetStateColor = "#AAAAAA"
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="101" then
			GetDivCodeName = "가맹점용 개별매입"
		elseif Fdivcode="111" then
			GetDivCodeName = "가맹점용 개별특정"
		elseif Fdivcode="121" then
			GetDivCodeName = "온라인특정->가맹점특정"
		elseif Fdivcode="131" then
			GetDivCodeName = "온라인특정->가맹점매입"
		elseif Fdivcode="201" then
			GetDivCodeName = "온라인매입->가맹점매입"
		elseif Fdivcode="251" then
			GetDivCodeName = "매입반품->오프재고"
		elseif Fdivcode="261" then
			GetDivCodeName = "오프재고->가맹점출고"
		elseif Fdivcode="300" then
			GetDivCodeName = "온라인주문"
		elseif Fdivcode="301" then
			GetDivCodeName = "온라인매입"
		elseif Fdivcode="302" then
			GetDivCodeName = "온라인특정"
		elseif Fdivcode="501" then
			GetDivCodeName = "직영샵주문"
		elseif Fdivcode="502" then
			GetDivCodeName = "수수료샵"
		elseif Fdivcode="503" then
			GetDivCodeName = "가맹점"
		else
			GetDivCodeName = ""
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="101" then
			GetDivCodeColor = "#0000AA"
		elseif Fdivcode="111" then
			GetDivCodeColor = "#AA0000"
		elseif Fdivcode="121" then
			GetDivCodeColor = "#AA00AA"
		elseif Fdivcode="131" then
			GetDivCodeColor = "#00AAAA"
		elseif Fdivcode="201" then
			GetDivCodeColor = "#AAAA00"
		elseif Fdivcode="300" then
			GetDivCodeColor = "#FF0000"
		elseif Fdivcode="501" then
			GetDivCodeColor = "#0000FF"
		elseif Fdivcode="502" then
			GetDivCodeColor = "#00FF00"
		elseif Fdivcode="503" then
			GetDivCodeColor = "#AAFFAA"
		else
			GetDivCodeColor = "#000000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CUpcheMwItem
	public FOnlineMwDiv
	public FOnlineDefaultmargine
	public FfranChargeDiv
	public FfranDefaultmargine

	public function IsOnOffWitak
		if (FOnlineMwDiv="W") and (FfranChargeDiv="2") then
			IsOnOffWitak = true
		else
			IsOnOffWitak = false
		end if
	end function

	public function GetOnlineMwDivName
		if FOnlineMwDiv="M" then
			GetOnlineMwDivName = "매입"
		elseif FOnlineMwDiv="W" then
			GetOnlineMwDivName = "특정"
		end if
	end function

	public function GetOnlineDefaultmargine
		GetOnlineDefaultmargine = FOnlineDefaultmargine
	end function

	public function GetfranChargeDivName
		if FfranChargeDiv="2" then
			GetfranChargeDivName = "특정"
		elseif FfranChargeDiv="4" then
			GetfranChargeDivName = "출고 정산"
		elseif FfranChargeDiv="5" then
			GetfranChargeDivName = "출고 정산"
		elseif FfranChargeDiv="6" then
			GetfranChargeDivName = "업체 특정"
		elseif FfranChargeDiv="8" then
			GetfranChargeDivName = "업체 매입"
		end if
	end function

	public function GefranDefaultmargine
		GefranDefaultmargine = FfranDefaultmargine
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

Class CUpcheMwInfo
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FRectdesignerId

	public Sub GetDesignerMWInfo()
		dim sqlStr, i
		sqlStr = "select top 1 maeipdiv,defaultmargine from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + FRectdesignerId + "'"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneItem = new CUpcheMwItem

		if Not rsget.Eof then
			FOneItem.FOnlineMwDiv = rsget("maeipdiv")
			FOneItem.FOnlineDefaultmargine = rsget("defaultmargine")
		end if
		rsget.Close

		sqlStr = "select chargediv,defaultmargin from [db_shop].[dbo].tbl_shop_designer"
		sqlStr = sqlStr + " where makerid='" + FRectdesignerId + "'"
		sqlStr = sqlStr + " and shopid='streetshop800'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			FOneItem.FfranChargeDiv = rsget("chargediv")
			FOneItem.FfranDefaultmargine = rsget("defaultmargin")
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class COrderSheetMasterItem
	public fsmssenddate
	public Ftplcompanyid
	public Fidx
	public Ftargetid
	public Fbaljuid
	public Freguser
	public Ffinishuser
	public Ftargetname
	public Fbaljuname
	public Fregname
	public Ffinishname
	public Fdivcode
	public Ftotalsellcash
	public Ftotalsuplycash
	public Ftotalbuycash
	public Ftotalreturnsellcash
	public Ftotalreturnsuplycash
	public Ftotalreturnbuycash
	public Fjumunsellcash
	public Fjumunsuplycash
	public Fjumunbuycash
	public Fvatinclude
	public Fregdate
	public Fupdt
	public Fdeldt
	public Fscheduledate
	public Fbeasongdate
	public Fipgodate
	public Fsongjangdiv
	public Fsongjangname
	public Fsongjangno
	public Fbaljucode
	public Fstatecd
	public Fcomment
	public FBrandList
	public FCount
	public FMakerid
	public Fscheduleipgodate
	public Freplycomment
	public Fsendsms
	public Fipkumdate
	public Fsegumdate
	public FaLinkCode
	public FbLinkCode
	public Fipchulsellcash
	public Fipchulsuplycash
	public Fipchulbuycash
	public Fipchuldeldt
	public Fjumunitemno
	public Ftotalitemno
	public Ftotalreturnitemno
	public FItemGubun
	public FItemId
	public FItemoption
	public FItemName
	public FItemOptionname
	public f10totalsellcash
	public f90totalsellcash
	public f70totalsellcash
	public f10totalsuplycash
	public f90totalsuplycash
	public f70totalsuplycash
	public Fbaljudate
	public Fbaljunum
	public fshopconfirmuserid
	public fshopconfirmdate
	public fshopconfirmipgodate
	public fcwflag
	public Fcheckusersn
	public Frackipgousersn
	public Fcheckusername
	public Frackipgousername
	public FcurrencyUnit
	public Fforeign_statecd
	public fsitename
	public fjumunforeign_sellcash
	public fjumunforeign_suplycash
	public ftotalforeign_sellcash
	public ftotalforeign_suplycash
	public fpurchasetype
	public FworkSecond
	public FfirstItemName
	public FcurrencyChar
	public FtotalDeliverPriceForeign
	public FfreightTerm
	public FopenState
	public FshippingAddress
	public FinvoiceAddress
	public finvoiceidx
	public Freportidx
	public Freportstate
    public FppMasterIdx
	public FppReportidx
	public FppReportstate
	public fmanager_hp
	public fmanager_email

	public function IsSendedSMS()
		if IsNULL(Fsendsms) then
			IsSendedSMS = false
		elseif Fsendsms="Y" then
			IsSendedSMS = true
		else
			IsSendedSMS = false
		end if
	end function

	public function getScheduleIpgodate()
		if IsNULL(Fscheduleipgodate) then
			getScheduleIpgodate = ""
		else
			getScheduleIpgodate = Left(Fscheduleipgodate,10)
		end if
	end function

	public function getScheduledate()
		if IsNULL(Fscheduledate) then
			getScheduledate = ""
		else
			getScheduledate = Left(Fscheduledate,10)
		end if
	end function

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

	public function GetJumunSellcashOrSellcash
		if IsFixed then
			GetJumunSellcashOrSellcash = Ftotalsellcash
		else
			GetJumunSellcashOrSellcash = Fjumunsellcash
		end if
	end function

	public function GetJumunBuycashOrBuycash
		if IsFixed then
			GetJumunBuycashOrBuycash = Ftotalbuycash
		else
			GetJumunBuycashOrBuycash = Fjumunbuycash
		end if
	end function

	public function GetJumunSuplycashOrSuplycash
		if IsFixed then
			GetJumunSuplycashOrSuplycash = Ftotalsuplycash
		else
			GetJumunSuplycashOrSuplycash = Fjumunsuplycash
		end if
	end function

	public function GetIpgoWhereStr()
		if Fdivcode="501" then
			GetIpgoWhereStr = ""
		else
			GetIpgoWhereStr = "경기도 포천시 군내면 용정경제로2길 83 텐바이텐 물류센터 "
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="101" then
			GetDivCodeName = "가맹점용 개별매입"
		elseif Fdivcode="111" then
			GetDivCodeName = "가맹점용 개별특정"
		elseif Fdivcode="121" then
			GetDivCodeName = "온라인특정->가맹점특정"
		elseif Fdivcode="131" then
			GetDivCodeName = "온라인특정->가맹점매입"
		elseif Fdivcode="201" then
			GetDivCodeName = "온라인매입->가맹점매입"
		elseif Fdivcode="251" then
			GetDivCodeName = "매입반품->오프재고"
		elseif Fdivcode="261" then
			GetDivCodeName = "오프재고->가맹점출고"
		elseif Fdivcode="300" then
			GetDivCodeName = "온라인주문"
		elseif Fdivcode="301" then
			GetDivCodeName = "온라인매입"
		elseif Fdivcode="302" then
			GetDivCodeName = "온라인특정"
		elseif Fdivcode="501" then
			GetDivCodeName = "직영샵주문"
		elseif Fdivcode="502" then
			GetDivCodeName = "수수료샵"
		elseif Fdivcode="503" then
			GetDivCodeName = "가맹점"
		else
			GetDivCodeName = ""
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="101" then
			GetDivCodeColor = "#0000AA"
		elseif Fdivcode="111" then
			GetDivCodeColor = "#AA0000"
		elseif Fdivcode="121" then
			GetDivCodeColor = "#AA00AA"
		elseif Fdivcode="131" then
			GetDivCodeColor = "#00AAAA"
		elseif Fdivcode="201" then
			GetDivCodeColor = "#AAAA00"
		elseif Fdivcode="300" then
			GetDivCodeColor = "#FF0000"
		elseif Fdivcode="501" then
			GetDivCodeColor = "#0000FF"
		elseif Fdivcode="502" then
			GetDivCodeColor = "#00FF00"
		elseif Fdivcode="503" then
			GetDivCodeColor = "#AAFFAA"
		else
			GetDivCodeColor = "#000000"
		end if
	end function

    public function getOrderpaymentstatus()
        dim buf : buf= ""
        'SHIPPING

        'if (Fstatecd=" ") then
		if (Fforeign_statecd=10) then
			buf = "결제완료"
		elseif (Fforeign_statecd=9) then
			buf = "입금대기"
		elseif (Fforeign_statecd=8) then
			buf = "결제대기"
		end if
        'end if

        if (buf="결제완료") then
            getOrderpaymentstatus = "<font color='#AA0000'>"&buf&"</font>"
        elseif (buf="입금대기") then
            getOrderpaymentstatus = "<font color='#0000AA'>"&buf&"</font>"
        elseif (buf="결제대기") then
			getOrderpaymentstatus = "<font color='#AAAAAA'>"&buf&"</font>"
        else
            getOrderpaymentstatus = buf
        end if
    end function

	public function GetStateName()
		if Fstatecd=" " and fforeign_statecd="0" then
			GetStateName = "업체접수(견적요청)"
		elseif Fstatecd=" " and fforeign_statecd="3" then
			GetStateName = "업체접수확인"
		elseif Fstatecd=" " and fforeign_statecd="7" then
			GetStateName = "업체접수완료"
		elseif Fstatecd=" " and (isnull(fforeign_statecd) or fforeign_statecd="") then
			GetStateName = "주문서작성중"
		elseif Fstatecd="0" then
			GetStateName = "주문접수"
		elseif Fstatecd="1" then
			GetStateName = "주문확인"
		elseif Fstatecd="2" then
			GetStateName = "입금대기"
		elseif Fstatecd="5" then
			GetStateName = "배송준비"
		elseif Fstatecd="6" then
			GetStateName = "출고대기"
		elseif Fstatecd="7" then
			GetStateName = "출고완료"
		elseif Fstatecd="8" then
			GetStateName = "검품완료<br>(입고대기)"
		elseif Fstatecd="9" then
			GetStateName = "입고완료"
		end if
	end function

	public function GetStateColor()
		if Fstatecd="0" then
			GetStateColor = "#00000"
		elseif Fstatecd="1" then
			GetStateColor = "#00AA00"
		elseif Fstatecd="2" then
			GetStateColor = "#0000AA"
		elseif Fstatecd="5" then
			GetStateColor = "#AAAA00"
		elseif Fstatecd="6" then
			GetStateColor = "#AA00AA"
		elseif Fstatecd="7" then
			GetStateColor = "#AA0000"
		elseif Fstatecd="8" then
			GetStateColor = "#33AAAA"
		elseif Fstatecd="9" then
			GetStateColor = "#AA33AA"
		elseif Fstatecd=" " then
			GetStateColor = "#AAAAAA"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLimitCheckDetailItem
	public Fidx
	public Fmasteridx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellcash
	public Fsuplycash
	public Fbuycash
	public Fbaljuitemno
	public Frealitemno
	public Fcomment
	public FMakerid
	public FSellYn
	public FDispYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public FIsNewItem

	'' Fcurrno = Frealstock + Fipkumdiv5 + Foffconfirmno
	public Fcurrno
	'' FMaystockno =
	public FMaystockno

	public function IsSoldOut
		IsSoldOut = ((FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa=0)))
	end function

	public function GetIsSlodOutText
		if IsSoldOut then
			GetIsSlodOutText = "Y"
		else
			GetIsSlodOutText = "N"
		end if
	end function

	public function GetIsSlodOutColor
		if IsSoldOut then
			GetIsSlodOutColor = "#FF2222"
		else
			GetIsSlodOutColor = "#000000"
		end if
	end function

	public function GetLimitEa
		GetLimitEa = FLimitNo-FLimitSold
		if GetLimitEa<1 then GetLimitEa=0
	end function

	public function GetMayCheckColor()
		if (Frealitemno<1) and (FSellYn="Y") then
			GetMayCheckColor = "#88CC88"
		elseif (Frealitemno>0) and (FSellYn="N") then
			GetMayCheckColor = "#8888CC"
		else
			GetMayCheckColor = "#FFFFFF"
		end if
	end function

	public function GetSellYnColor()
		if FSellYn="N" then
			GetSellYnColor = "#FF2222"
		else
			GetSellYnColor = "#000000"
		end if
	end function

	public function GetDispYnColor()
		if FDispYn="N" then
			GetDispYnColor = "#FF2222"
		else
			GetDispYnColor = "#000000"
		end if
	end function

	public function GetLimitYnColor()
		if FLimitYn="Y" then
			GetLimitYnColor = "#FF2222"
		else
			GetLimitYnColor = "#000000"
		end if
	end function

	public function GetRealOrJumunNo()
		if FRectIsFixed then
			GetRealOrJumunNo = Frealitemno
		else
			GetRealOrJumunNo = Fbaljuitemno
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COrderSheetDetailItem
	public Fstatecd
	public fsocname
	public Flcitemname
	public Flcitemoptionname
	public flcprice
	public fforeign_sellcash
	public fforeign_suplycash
	public Fidx
	public Fmasteridx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	Public Forgsellprice
	public Fsellcash
	public Fsuplycash
	public Fbuycash
	public Fbaljuitemno
	Public Frealbaljuitemno
	public Frealitemno
	public Fcheckitemno
	public Fregdate
	public Fupdt
	public Fdeldt
	public Fbaljudiv
	public Fcomment
	public FMakerid
	public FItemDefaultMwDiv
	public Fdeliverytype
	public FRectIsFixed
	public Fonlinesellcash
	public Fonlinebuycash
	public FoffChargeDiv
	public Fshopdefaultmargin
	public Fshopdefaultsuplymargin
	public FIpgoFlag
	public Fdefaultmaginflag
	public Fbuymaginflag
	public Fsuplymaginflag
	public Fprtidx
	public FonlineSellyn
	public FonlineDispyn
	public FonlineLimityn
	public FonlineLimitno
	public FonlineLimitsold
	public Fsmallimage
	public FOffimgMain
	public FOffimgList
	public FOffimgSmall
	public FPublicBarcode
	public Fdetail_status
	public Fdetail_description
    public Fcentermwdiv
    public Fboxsongjangno
    public FUpcheManageCode
	public FcurrencyUnit
	public Fbasicimage
	public Fbaljucode
	public fproductidx
	public fproductidxArr
	public Freportidx
	public FScheduleDate
	public fstatecdname
	public fblinkcode
	public ftenbarcode
	public fbarcode
	public Fmwdiv
	public fpurchaseTypename
	public fcateName
	public flastIpgoDate

    public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = Fsmallimage
		else
			GetImageSmall = Foffimgsmall
		end if
	end function

	public function GetLimitEa()
		GetLimitEa = FonlineLimitno-FonlineLimitsold
		if GetLimitEa<1 then GetLimitEa=0
	end function

	public function GetOnlineBigoStr()
		if (FonlineSellyn="N") or (FonlineDispyn="N") then
			GetOnlineBigoStr = "품절"
		end if

		if FonlineLimityn="Y" then
			GetOnlineBigoStr = GetOnlineBigoStr + " 한정(" + CStr(GetLimitEa()) + ")"
		end if
	end function

	public function GetNoinputDefaultmaginflag()
		if IsUpchebeasong then
			GetNoinputDefaultmaginflag = "U"
		else
			GetNoinputDefaultmaginflag = FItemDefaultMwDiv
		end if
	end function

	public function GetNoinputBuymaginflag()
		GetNoinputBuymaginflag = FoffChargeDiv
	end function

	public function GetNoinputSuplymaginflag()
		GetNoinputSuplymaginflag = FoffChargeDiv
	end function

	public function IsWi2Meaip()
		IsWi2Meaip = (IsOnLineItem) and (not IsUpchebeasong) and (FItemDefaultMwDiv="W") and (FoffChargeDiv="4")
	end function

	public function IsOnLineItem()
		IsOnLineItem = (Fitemgubun="10")
	end function

    ''수정요망.
	public function GetOn2Off2DivName()
		if IsUpchebeasong then
			if FoffChargeDiv="4" then     ''매입
				GetOn2Off2DivName = "<font color=#4444CC>업</font>" + CStr(GetOnlineMargine) + "매" + CStr(Getshopdefaultmargin) + "매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="5" then     ''출고정산 : 업체배송인경우 구분요망 - 오프라인 매입구분으로
				GetOn2Off2DivName = "<font color=#4444CC>업</font>" + CStr(GetOnlineMargine) + "위" + CStr(Getshopdefaultmargin) + "매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="2" then     ''매입
				GetOn2Off2DivName = "<font color=#4444CC>업</font>" + CStr(GetOnlineMargine) + "<font color=#AAAAAA>위</font>" + CStr(Getshopdefaultmargin) + "매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="6" then  ''업체특정
				GetOn2Off2DivName = "<font color=red>" + "업" + CStr(GetOnlineMargine) + "업위" + CStr(Getshopdefaultsuplymargin) + "</font>"
			elseif FoffChargeDiv="8" then  ''업체매입
				GetOn2Off2DivName = "<font color=red>" + "업" + CStr(GetOnlineMargine) + "업매" + CStr(Getshopdefaultsuplymargin)  + "</font>"
			else
				GetOn2Off2DivName = "<font color=red>" + "업" + CStr(GetOnlineMargine) + "<font color=red>???</font>"
			end if
		elseif FItemDefaultMwDiv="M" then
			if FoffChargeDiv="4" then
				if (GetOnlineMargine<>Getshopdefaultmargin) then
					GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "<font color=red>매" + CStr(Getshopdefaultmargin) + "</font>매" + CStr(Getshopdefaultsuplymargin)
				else
					GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "매" + CStr(Getshopdefaultsuplymargin)
				end if
			elseif FoffChargeDiv="5" then     ''출고정산
				GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "매" + CStr(Getshopdefaultsuplymargin)

			elseif FoffChargeDiv="2" then
				GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "<font color=red>위" + CStr(Getshopdefaultmargin) + "</font>"
			elseif FoffChargeDiv="6" then  ''업체특정
				GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "<font color=red>업위" + CStr(Getshopdefaultmargin) + "</font>"
			elseif FoffChargeDiv="8" then  ''업체매입
				GetOn2Off2DivName = "매" + CStr(GetOnlineMargine) + "<font color=red>업매" + CStr(Getshopdefaultmargin) + "</font>"
			end if
		elseif FItemDefaultMwDiv="W" then
			if FoffChargeDiv="4" then     ''매입
				GetOn2Off2DivName = "<font color=#AAAAAA>위</font>" + CStr(GetOnlineMargine) + "<font color=red>매" + CStr(Getshopdefaultmargin) + "</font>매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="5" then     ''출고정산
				GetOn2Off2DivName = "<font color=#AAAAAA>위</font>" + CStr(GetOnlineMargine) + "<font color=red>매" + CStr(Getshopdefaultmargin) + "</font>매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="2" then  ''특정
				GetOn2Off2DivName = "<font color=#AAAAAA>위</font>" + CStr(GetOnlineMargine) + "<font color=#AAAAAA>위</font>" + CStr(Getshopdefaultmargin) + "<font color=#AAAAAA>위</font>" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="6" then  ''업체특정
				GetOn2Off2DivName = "<font color=red>" + "위" + CStr(GetOnlineMargine) + "업위" + CStr(Getshopdefaultsuplymargin) + "</font>"
			elseif FoffChargeDiv="8" then  ''업체매입
				GetOn2Off2DivName = "<font color=red>" + "위" + CStr(GetOnlineMargine) + "업매" + CStr(Getshopdefaultsuplymargin)  + "</font>"
			else
				GetOn2Off2DivName = "<font color=red>" + "위" + CStr(GetOnlineMargine) + "<font color=red>???</font>"
			end if
		else
			if FoffChargeDiv="4" then     ''매입
				GetOn2Off2DivName = "??<font color=red>매" + CStr(Getshopdefaultmargin) + "</font>매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="5" then     ''출고정산
				GetOn2Off2DivName = "??<font color=red>매" + CStr(Getshopdefaultmargin) + "</font>매" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="2" then  ''특정
				GetOn2Off2DivName = "??<font color=#AAAAAA>위</font>" + CStr(Getshopdefaultmargin) + "<font color=#AAAAAA>위</font>" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="6" then  ''업체특정
				GetOn2Off2DivName = "??업위" + CStr(Getshopdefaultsuplymargin) + "</font>"
			elseif FoffChargeDiv="8" then  ''업체매입
				GetOn2Off2DivName = "??업매" + CStr(Getshopdefaultsuplymargin)  + "</font>"
			else
				GetOn2Off2DivName = "<font color=red>???-???</font>"
			end if
		end if
	end function

	public function getonOffDiffMaginName()
		if FItemDefaultMwDiv="M" then
			if FoffChargeDiv<>"4" then
				getonOffDiffMaginName = "매:" + CStr(GetOnlineMargine) + "->" + CStr(GetoffChargeDivName) + ":" + CStr(GetOfflineSuplyMargine)
			elseif FoffChargeDiv="4" then
				if (GetOnlineMargine<>Getshopdefaultmargin) or (GetOfflineSuplyMargine<>Getshopdefaultsuplymargin) then
					getonOffDiffMaginName = "매:" + CStr(Getshopdefaultmargin) + "->" + CStr(GetoffChargeDivName) + ":" + CStr(Getshopdefaultsuplymargin)
				end if
			else
				getonOffDiffMaginName = "매:" + CStr(Getshopdefaultmargin) + "->" + CStr(GetoffChargeDivName) + ":" + CStr(Getshopdefaultsuplymargin)
			end if
		else
			if FoffChargeDiv="4" then     ''매입
				'if (GetOnlineMargine<>Getshopdefaultmargin) or (GetOfflineSuplyMargine<>Getshopdefaultsuplymargin) then
					getonOffDiffMaginName = "위:" + CStr(GetOnlineMargine) + "->" + "매:" + CStr(Getshopdefaultmargin) + ")->매:" + CStr(Getshopdefaultsuplymargin)
				'end if
			elseif FoffChargeDiv="2" then  ''특정
				getonOffDiffMaginName = "위:" + CStr(GetOnlineMargine) + "->" + "위:" + CStr(Getshopdefaultmargin) + ")->매:" + CStr(Getshopdefaultsuplymargin)
			elseif FoffChargeDiv="6" then  ''업체특정
				getonOffDiffMaginName = "<font color=red>" + "위:" + CStr(GetOnlineMargine) + "->" + "업체특정:" + CStr(Getshopdefaultsuplymargin) + "</font>"
			elseif FoffChargeDiv="8" then  ''업체매입
				getonOffDiffMaginName = "<font color=red>" + "위:" + CStr(GetOnlineMargine) + "->" + "업체매입:" + CStr(Getshopdefaultsuplymargin)  + "</font>"
			else
				getonOffDiffMaginName = "<font color=red>???</font>"
			end if
		end if
	end function

	public function getChulgoDivName()
		if FItemDefaultMwDiv="M" then
			getChulgoDivName = "매(" + CStr(GetOnlineMargine) + ")->매(" + CStr(GetOfflineSuplyMargine) + ")"
		else
			if FoffChargeDiv="2" then
				getChulgoDivName = "위(" + CStr(GetOnlineMargine) + ")->위(" + CStr(GetOfflineSuplyMargine) + ")"
			elseif FoffChargeDiv="4" then
				getChulgoDivName = "위(" + CStr(GetOnlineMargine) + ")->매(" + CStr(GetOfflineSuplyMargine) + ")"
			end if
		end if
	end function

	public function GetoffChargeDivName
		if FoffChargeDiv="2" then
			GetoffChargeDivName = "특정"
		elseif FoffChargeDiv="4" then
			GetoffChargeDivName = "매입"
		elseif FoffChargeDiv="6" then
			GetoffChargeDivName = "업체 특정"
		elseif FoffChargeDiv="8" then
			GetoffChargeDivName = "업체 매입"
		else
			GetoffChargeDivName = FoffChargeDiv
		end if
	end function

	public function GetOrgShopSuplycashbyMargine()
		GetOrgShopSuplycashbyMargine = Fsellcash - CLng(Fsellcash*Getshopdefaultsuplymargin/100)
	end function

	public function Getshopdefaultmargin
		Getshopdefaultmargin = Fshopdefaultmargin
	end function

	public function Getshopdefaultsuplymargin
		Getshopdefaultsuplymargin = Fshopdefaultsuplymargin
	end function

	public function GetOnlineMargine()
		if Fonlinesellcash<>0 then
			GetOnlineMargine = (100-CLng(Fonlinebuycash/Fonlinesellcash*100*100)/100)
		end if
	end function

	public function GetOfflineSuplyMargine()
		if Fbuycash<>0 then
			GetOfflineSuplyMargine = (100-CLng(Fsuplycash/Fsellcash*100*100)/100)
		end if
	end function

	public function IsUpchebeasong()
		if (Fdeliverytype="2") or (Fdeliverytype="9") or (Fdeliverytype="7") then
			IsUpchebeasong = true
		else
			IsUpchebeasong = false
		end if
	end function

	public function GetMwDivColor()
		if IsUpchebeasong then
			GetMwDivColor = "#EE4444"
		elseif FItemDefaultMwDiv="M" then
			GetMwDivColor = "#4444EE"
		else
			GetMwDivColor = "#000000"
		end if
	end function

	public function GetRealOrJumunNo()
		if FRectIsFixed then
			GetRealOrJumunNo = Frealitemno
		else
			GetRealOrJumunNo = Fbaljuitemno
		end if
	end function


	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COrderSheet
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public fArrLIst
	public FRectIdx
	public FRectDivCode
	public FRectStatecd
	public FRectBaljuId
	public FRectBaljuname
	public FRectbaljunum
	public FRectTargetid
	public FRectTargetName
	public FRectReguser
	public FRectRegname
	public FRectScheduledate
	public FRectStartDate
	public FRectEndDate
	public FRectIsFixed
	public FRectMakerid
	public FRectIdxArr
	public FRectDivCodeArr
	public FRectDivCodeUnder
	public FRectStateCdOver
	public FRectStateCdOver2
	public FRectBaljuCode
	public FRectShopid
	public FRectNotIpgoOnly
	public FRectMinusOnly
	public FRectShopDiv
	public FRectDivGubun
    public FRectReOrderOnly
	public FRectComment
	public FRectFromDate
	public FRectToDate
	public frectdatefg
	public frectitemgubun
	public FRectitemid
	public FRectitemoption
	public FRectShortYN
	public FRectIncludePreOrderNo
	public FRectmwdiv
	public FRectITemGubunArr
	public FRectItemIdArr
	public FRectItemOptionArr
	public FRectOrderNoArr
	public FRectdatetype
	public FRectOrgBaljuCode
    public FRectBrandPurchaseType
	public FRectpurchasetype
	public FRectBLinkCode
	public FtplGubun
	public FRecttplgubun
	public FRectproductidx
	Public FGroupByBaljuCode
	public frecttotalyn
	Public total_jumunsellcash
	Public total_jumunsuplycash
	Public total_totalsuplycash

	public FRectReportState
	public FRectSearchField
	public FRectSearchText
	public FAverageWorkSecond

	public Sub GetOrderSheetMasterByBrandSum()
		dim i,sqlStr
		sqlStr = " select m.idx, m.baljucode, m.baljuid, convert(varchar(19),m.regdate,20) as regdate, m.scheduledate,"
		sqlStr = sqlStr + " m.baljuname, m.statecd, m.ipgodate, m.divcode, sum(d.sellcash*d.realitemno) as selltotal,"
		sqlStr = sqlStr + " sum(d.buycash*d.realitemno) as buytotal,"
		sqlStr = sqlStr + " sum(d.realitemno) as counttotal"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and m.idx=d.masteridx"

		if FRectDivCodeUnder<>"" then
			sqlStr = sqlStr + " and m.divcode<" + FRectDivCodeUnder + ""
		end if

		if FRectStateCdOver<>"" then
			sqlStr = sqlStr + " and m.statecd>='" + FRectStateCdOver + "'"
		end if

		if FRectDivCodeArr<>"" then
			sqlStr = sqlStr + " and m.divcode in " + FRectDivCodeArr + ""
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		sqlStr = sqlStr + " group by m.idx, m.baljucode, m.baljuid,"
		sqlStr = sqlStr + " m.regdate, m.scheduledate, m.baljuname, m.statecd, m.ipgodate, m.divcode"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COrderSheetCheckItem
				FItemList(i).Fidx       = rsget("idx")
				FItemList(i).Fbaljuid   = rsget("baljuid")
				FItemList(i).Fbaljucode     = rsget("baljucode")
				FItemList(i).Ftotalsellcash = rsget("selltotal")
				FItemList(i).Ftotalbuycash  = rsget("buytotal")
				FItemList(i).Fcount         = rsget("counttotal")
				FItemList(i).FRegDate         = rsget("regdate")
				FItemList(i).FScheduledate         = rsget("scheduledate")
				FItemList(i).FBaljuName         = db2html(rsget("baljuname"))
				FItemList(i).Fstatecd		= rsget("statecd")
				FItemList(i).Fipgodate		= rsget("ipgodate")
				FItemList(i).Fdivcode	= rsget("divcode")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public Sub GetOrderSheetByBrandSum()
		dim i,sqlStr
		sqlStr = " select A.*,"
		sqlStr = sqlStr + " IsNULL(B.selltotal2,0) as selltotal2, IsNULL(B.buytotal2,0) as buytotal2,"
		sqlStr = sqlStr + " IsNULL(B.counttotal2,0) as counttotal2"

		sqlStr = sqlStr + " from (select d.makerid, sum(d.sellcash*d.realitemno) as selltotal,"
		sqlStr = sqlStr + " sum(d.buycash*d.realitemno) as buytotal,"
		sqlStr = sqlStr + " sum(d.realitemno) as counttotal"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and m.idx=d.masteridx"
		if FRectBaljuid<>"" then
			sqlStr = sqlStr + " and m.baljuid='" + FRectBaljuid + "'"
		end if

		if FRectStateCdOver<>"" then
			sqlStr = sqlStr + " and m.statecd>='" + FRectStateCdOver + "'"
		end if

		if FRectStatecd<>"" then
			sqlStr = sqlStr + " and m.statecd='" + FRectStatecd + "'"
		end if

		if FRectDivCodeArr<>"" then
			sqlStr = sqlStr + " and m.divcode in " + FRectDivCodeArr + ""
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		sqlStr = sqlStr + " group by d.makerid"
		sqlStr = sqlStr + " ) as A"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " (select d.makerid, sum(d.sellcash*d.realitemno) as selltotal2,"
			sqlStr = sqlStr + " sum(d.buycash*d.realitemno) as buytotal2,"
			sqlStr = sqlStr + " sum(d.realitemno) as counttotal2"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
			sqlStr = sqlStr + " where m.deldt is NULL"
			sqlStr = sqlStr + " and d.deldt is NULL"
			sqlStr = sqlStr + " and m.idx=d.masteridx"

			if FRectStateCdOver<>"" then
				sqlStr = sqlStr + " and m.statecd>='" + FRectStateCdOver2 + "'"
			end if

			if FRectDivCodeUnder<>"" then
				sqlStr = sqlStr + " and m.divcode<" + FRectDivCodeUnder + ""
			end if

			if FRectMakerid<>"" then
				sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
			end if
			sqlStr = sqlStr + " group by d.makerid"
			sqlStr = sqlStr + " ) as B"
			sqlStr = sqlStr + " on A.makerid=B.makerid"
		sqlStr = sqlStr + " order by A.makerid"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COrderSheetCheckItem
				FItemList(i).FMakerid       = rsget("makerid")
				FItemList(i).Ftotalsellcash = rsget("selltotal")
				FItemList(i).Ftotalbuycash  = rsget("buytotal")
				FItemList(i).Fcount         = rsget("counttotal")
				FItemList(i).Ftotalsellcash2 = rsget("selltotal2")
				FItemList(i).Ftotalbuycash2  = rsget("buytotal2")
				FItemList(i).Fcount2         = rsget("counttotal2")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetFranBalju2UpcheBaljuBrandlist()
		dim i,sqlStr
		sqlStr = "select d.makerid, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " , d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " , sum(Case when m.statecd='0' then d.baljuitemno else 0 end ) as jupsucnt"
		sqlStr = sqlStr + " , sum(Case when m.statecd='1' then d.baljuitemno else 0 end ) as cnt"
		sqlStr = sqlStr + " , IsNULL(s.realstock,0) as realstock"
		sqlStr = sqlStr + " , i.mwdiv "
		sqlStr = sqlStr + " , IsNULL(s.preorderno,0) as preorderno "
		sqlStr = sqlStr + " , IsNULL(s.preordernofix,0) as preordernofix "
		sqlStr = sqlstr + " , T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on (d.itemgubun='10' and d.itemid=i.itemid )"
        sqlStr = sqlStr + "     left join db_summary.dbo.tbl_current_logisstock_summary s"
        sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_Stock T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=T.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=T.itemid "
		sqlStr = sqlStr + " 		and d.itemoption=T.itemoption "

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and m.targetid='" + FRectBaljuId + "'"
		end if

		if (FRectStatecd<>"") then
			sqlStr = sqlStr + " and m.statecd='" + FRectStatecd + "'"
		else
			sqlStr = sqlStr + " and m.statecd in ('0','1')"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		sqlStr = sqlStr + " and d.baljuitemno>0"
		sqlStr = sqlStr + " and ((d.itemgubun<>'10') or (i.mwdiv='U'))"		'업배 또는 오프상품만
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " group by d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, IsNULL(s.realstock,0), i.mwdiv, IsNULL(s.preorderno,0), IsNULL(s.preordernofix,0), T.StockReipgoDate "

		if (FRectShortYN = "Y") or (FRectIncludePreOrderNo = "Y") then
			if FRectIncludePreOrderNo = "Y" then
				sqlStr = sqlStr + " having ((IsNULL(s.realstock,0) - sum(Case when m.statecd in ('0', '1') then d.baljuitemno else 0 end ) + IsNULL(s.preorderno,0)) < 0) "
			else
				sqlStr = sqlStr + " having ((IsNULL(s.realstock,0) - sum(Case when m.statecd in ('0', '1') then d.baljuitemno else 0 end )) < 0) "
			end if
		end if

		sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid"
		'response.write sqlStr

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljusumItem
				FItemList(i).FMakerid          	= rsget("makerid")
				FItemList(i).FItemGubun        	= rsget("itemgubun")
				FItemList(i).FItemId           	= rsget("itemid")
				FItemList(i).FItemoption       	= rsget("itemoption")
				FItemList(i).FJupsuCount       	= rsget("jupsucnt")
				FItemList(i).FCount            	= rsget("cnt")
				FItemList(i).FItemName         	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionname   	= db2html(rsget("itemoptionname"))
                FItemList(i).Frealstock        	= rsget("realstock")

                FItemList(i).Fmwdiv				= rsget("mwdiv")
                FItemList(i).Fpreorderno        = rsget("preorderno")
                FItemList(i).Fpreordernofix     = rsget("preordernofix")
                FItemList(i).FreipgoMayDate     = rsget("reipgoMayDate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetFranBalju2UpcheBaljuBrandlistNew_20111102()
		dim i,sqlStr
		sqlStr = "select d.makerid, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " , d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " , (Case when m.statecd='0' then d.baljuitemno else 0 end ) as jupsucnt"
		sqlStr = sqlStr + " , (Case when m.statecd='1' then d.baljuitemno else 0 end ) as cnt"
		sqlStr = sqlStr + " , IsNULL(s.realstock,0) as realstock"
		sqlStr = sqlStr + " , i.mwdiv "
		sqlStr = sqlStr + " , IsNULL(s.preorderno,0) as preorderno "
		sqlStr = sqlStr + " , IsNULL(s.preordernofix,0) as preordernofix "
		sqlStr = sqlstr + " , T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlStr + " , m.baljuid as shopid "
		sqlStr = sqlStr + " , m.baljuname as shopname "
		sqlStr = sqlStr + " , m.baljucode "
		sqlStr = sqlStr + " , d.upcheorderlinkcode "

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		m.idx=d.masteridx "
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on (d.itemgubun='10' and d.itemid=i.itemid )"

        sqlStr = sqlStr + "     left join db_summary.dbo.tbl_current_logisstock_summary s"
        sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_Stock T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=T.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=T.itemid "
		sqlStr = sqlStr + " 		and d.itemoption=T.itemoption "

		sqlStr = sqlStr + " where 1 = 1 "
		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and m.targetid='" + FRectBaljuId + "'"
		end if

		if (FRectStatecd<>"") then
			sqlStr = sqlStr + " and m.statecd='" + FRectStatecd + "'"
		else
			sqlStr = sqlStr + " and m.statecd in ('0','1')"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		sqlStr = sqlStr + " and d.baljuitemno>0"
		sqlStr = sqlStr + " and ((d.itemgubun<>'10') or (i.mwdiv='U'))"		'업배 또는 오프상품만
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		'sqlStr = sqlStr + " group by d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, IsNULL(s.realstock,0), i.mwdiv, IsNULL(s.preorderno,0), IsNULL(s.preordernofix,0), T.StockReipgoDate "

		if (FRectShortYN = "Y") or (FRectIncludePreOrderNo = "Y") then
			if FRectIncludePreOrderNo = "Y" then
				'sqlStr = sqlStr + " having ((IsNULL(s.realstock,0) - sum(Case when m.statecd in ('0', '1') then d.baljuitemno else 0 end ) + IsNULL(s.preorderno,0)) < 0) "
			else
				'sqlStr = sqlStr + " having ((IsNULL(s.realstock,0) - sum(Case when m.statecd in ('0', '1') then d.baljuitemno else 0 end )) < 0) "
			end if
		end if

		sqlStr = sqlStr + " order by d.makerid, m.baljuid, m.baljucode, d.itemgubun, d.itemid"
		'response.write sqlStr

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljusumItem
				FItemList(i).FMakerid          	= rsget("makerid")
				FItemList(i).FItemGubun        	= rsget("itemgubun")
				FItemList(i).FItemId           	= rsget("itemid")
				FItemList(i).FItemoption       	= rsget("itemoption")
				FItemList(i).FJupsuCount       	= rsget("jupsucnt")
				FItemList(i).FCount            	= rsget("cnt")
				FItemList(i).FItemName         	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionname   	= db2html(rsget("itemoptionname"))
                FItemList(i).Frealstock        	= rsget("realstock")

                FItemList(i).Fmwdiv				= rsget("mwdiv")
                FItemList(i).Fpreorderno        = rsget("preorderno")
                FItemList(i).Fpreordernofix     = rsget("preordernofix")
                FItemList(i).FreipgoMayDate     = rsget("reipgoMayDate")

                FItemList(i).Fshopid				= rsget("shopid")
                FItemList(i).Fshopname				= db2html(rsget("shopname"))
                FItemList(i).Fbaljucode				= rsget("baljucode")

                FItemList(i).Fupcheorderlinkcode	= rsget("upcheorderlinkcode")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetFranBalju2UpcheBaljuBrandlistNewProc()
		dim i,sqlStr

		sqlStr = " exec [db_storage].[dbo].[sp_Ten_UpcheOrderList] '" & FRectMakerid & "', '" & FRectBaljuId & "', '" & FRectStartDate & "', '" & FRectShortYN & "', '" & FRectIncludePreOrderNo & "' "
		''response.write sqlStr & "<br><br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljusumItem
				FItemList(i).FMakerid          	= rsget("makerid")
				FItemList(i).FItemGubun        	= rsget("itemgubun")
				FItemList(i).FItemId           	= rsget("itemid")
				FItemList(i).FItemoption       	= rsget("itemoption")
				FItemList(i).FJupsuCount       	= rsget("jupsucnt")
				FItemList(i).FCount            	= rsget("cnt")
				FItemList(i).FItemName         	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionname   	= db2html(rsget("itemoptionname"))
                FItemList(i).Frealstock        	= rsget("realstock")

                FItemList(i).Fmwdiv				= rsget("mwdiv")
                FItemList(i).Fpreorderno        = rsget("preorderno")		''// + (rsget("preorderno9") - rsget("preordernofix9"))
                FItemList(i).Fpreordernofix     = rsget("preordernofix")
                FItemList(i).FreipgoMayDate     = rsget("reipgoMayDate")

                FItemList(i).Fshopid				= rsget("shopid")
                FItemList(i).Fshopname				= db2html(rsget("shopname"))
                FItemList(i).Fbaljucode				= rsget("baljucode")
				FItemList(i).Fbaljucodecnt			= rsget("baljucodecnt")

                FItemList(i).Fupcheorderlinkcode	= rsget("upcheorderlinkcode")

                FItemList(i).Fdeliverytype	= rsget("deliverytype")
                FItemList(i).Fcentermwdiv	= rsget("centermwdiv")

				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("offimgsmall")

				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""

				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall

				FItemList(i).FpriceCnt = rsget("priceCnt")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetFranBalju2UpcheBaljuBrandlistNewProcNEW()
		dim i,sqlStr

		sqlStr = " exec [db_storage].[dbo].[sp_Ten_UpcheOrderListNEW] '" & FRectMakerid & "', '" & FRectBaljuId & "', '" & FRectStartDate & "', '" & FRectShortYN & "', '" & FRectIncludePreOrderNo & "', '" & FGroupByBaljuCode & "' "
		response.write sqlStr & "<br><br>"
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljusumItem
				FItemList(i).FMakerid          	= rsget("makerid")
				FItemList(i).FItemGubun        	= rsget("itemgubun")
				FItemList(i).FItemId           	= rsget("itemid")
				FItemList(i).FItemoption       	= rsget("itemoption")
				FItemList(i).FJupsuCount       	= rsget("jupsucnt")
				FItemList(i).FCount            	= rsget("cnt")
				FItemList(i).FItemName         	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionname   	= db2html(rsget("itemoptionname"))

                FItemList(i).Frealstock        	= rsget("realstock")
				if (FItemList(i).Frealstock < 0) then
					FItemList(i).Frealstock = 0
				end if

                FItemList(i).Fmwdiv				= rsget("mwdiv")
                FItemList(i).Fpreorderno        = rsget("preorderno")		''// + (rsget("preorderno9") - rsget("preordernofix9"))
                FItemList(i).Fpreordernofix     = rsget("preordernofix")
                FItemList(i).FreipgoMayDate     = rsget("reipgoMayDate")

                FItemList(i).Fshopid				= rsget("shopid")
                FItemList(i).Fshopname				= db2html(rsget("shopname"))
                FItemList(i).Fbaljucode				= rsget("baljucode")
				FItemList(i).Fbaljucodecnt			= rsget("baljucodecnt")

                FItemList(i).Fupcheorderlinkcode	= rsget("upcheorderlinkcode")

                FItemList(i).Fdeliverytype	= rsget("deliverytype")
                FItemList(i).Fcentermwdiv	= rsget("centermwdiv")

				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("offimgsmall")

				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""

				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall

				FItemList(i).FpriceCnt = rsget("priceCnt")
				FItemList(i).FpreUnderCnt = rsget("preUnderCnt")
				FItemList(i).Fonbaljuitemno = rsget("onbaljuitemno")

				FItemList(i).Ftotbaljuitemno = rsget("totbaljuitemno")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetFranBalju2UpcheBaljuBrandlistNew()
		dim i,sqlStr
		sqlStr = "select d.makerid, d.itemgubun, d.itemid, d.itemoption, i.smallimage, si.offimgsmall, i.deliverytype, si.centermwdiv "
		sqlStr = sqlStr + " , d.itemname, d.itemoptionname"

		sqlStr = sqlStr + " , IsNULL(s.realstock,0) as realstock"
		sqlStr = sqlStr + " , i.mwdiv "
		sqlStr = sqlStr + " , sum(Case when m.statecd='0' then d.baljuitemno when m.statecd ='1' then d.baljuitemno - d.realbaljuitemno else 0 end ) as jupsucnt"
		sqlStr = sqlStr + " , sum(Case when m.statecd = '1' then d.realbaljuitemno when m.statecd = '6' then d.realitemno else 0 end ) as cnt"
		'sqlStr = sqlStr + " , IsNULL(s.offjupno,0) as jupsucnt "
		'sqlStr = sqlStr + " , 0 as cnt"

		sqlStr = sqlStr + " , IsNULL(s.preorderno,0) as preorderno "
		sqlStr = sqlStr + " , IsNULL(s.preordernofix,0) as preordernofix "

		'입고완료된 기주문 수량(중복주문 안한다)
		sqlStr = sqlStr + " , IsNull(sum(distinct  Case when  ujm.baljucode is not null and ujm.statecd = '9' then ujd.baljuitemno else 0 end ), 0) as preorderno9 "
		sqlStr = sqlStr + " , IsNull(sum(distinct  Case when  ujm.baljucode is not null and ujm.statecd = '9' then ujd.realitemno else 0 end ), 0) as preordernofix9 "

		sqlStr = sqlstr + " , T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlStr + " , '' as shopid "
		''sqlStr = sqlStr + " , '' as shopname "

		sqlStr = sqlStr + " , (case when count(distinct m.baljuname) > 1 then max(m.baljuname) + ' 외 ' + cast((count(distinct m.baljuname) - 1) as varchar(32)) else max(m.baljuname) end) as shopname "

		If FGroupByBaljuCode = True Then
			sqlStr = sqlStr + " , m.baljucode as baljucode "
		Else
			sqlStr = sqlStr + " , max(m.baljucode) as baljucode "
		End If

		sqlStr = sqlStr + " , count(m.baljucode) as baljucodecnt "

		sqlStr = sqlStr + " , (case when count(m.baljucode) > 1 then max(m.baljucode) + ' 외 ' + cast((count(m.baljucode) - 1) as varchar(32)) else max(m.baljucode) end) as upcheorderlinkcode "
		sqlStr = sqlStr + " , count(distinct d.buycash) as priceCnt "

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		m.idx=d.masteridx "

		sqlStr = sqlStr + " 	left join [db_storage].[dbo].tbl_ordersheet_master ujm "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.upcheorderlinkcode = ujm.baljucode "
		sqlStr = sqlStr + " 		and ujm.deldt is null "

		sqlStr = sqlStr + " 	left join [db_storage].[dbo].tbl_ordersheet_detail ujd "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and ujm.idx=ujd.masteridx "
		sqlStr = sqlStr + " 		and d.itemgubun=ujd.itemgubun and d.itemid=ujd.itemid and d.itemoption=ujd.itemoption  "
		sqlStr = sqlStr + " 		and ujd.deldt is null "

		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on (d.itemgubun='10' and d.itemid=i.itemid )"

		sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item si "
		sqlStr = sqlStr + "     on d.itemgubun=si.itemgubun and d.itemid=si.shopitemid and d.itemoption=si.itemoption "

        sqlStr = sqlStr + "     left join db_summary.dbo.tbl_current_logisstock_summary s"
        sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_Stock T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=T.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=T.itemid "
		sqlStr = sqlStr + " 		and d.itemoption=T.itemoption "

		sqlStr = sqlStr + " where 1 = 1 "
		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and m.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectTargetid<>"" then
			sqlStr = sqlStr + " and m.targetid='" + FRectTargetid + "'"
		end if

		if (FRectStatecd<>"") then
			sqlStr = sqlStr + " and m.statecd='" + FRectStatecd + "'"
		else
			sqlStr = sqlStr + " and m.statecd in ('0','1', '6')"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and m.regdate>='" + FRectStartDate + "'"
			''sqlStr = sqlStr + " and m.regdate_YYYYMMDD>='" + FRectStartDate + "'"
		end if

		sqlStr = sqlStr + " and d.baljuitemno>0"
		sqlStr = sqlStr + " and ((d.itemgubun<>'10') or (i.mwdiv='U'))"		'업배 또는 오프상품만
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " group by d.makerid, d.itemgubun, d.itemid, d.itemoption, i.smallimage, si.offimgsmall, i.deliverytype, si.centermwdiv, d.itemname, d.itemoptionname, IsNULL(s.realstock,0), i.mwdiv, IsNULL(s.offjupno,0), IsNULL(s.preorderno,0), IsNULL(s.preordernofix,0), T.StockReipgoDate "
		If FGroupByBaljuCode = True Then
			sqlStr = sqlStr + ",m.baljucode"
		End If

		if ((FRectShortYN = "Y") or (FRectIncludePreOrderNo = "Y")) and Not (FGroupByBaljuCode = True) then
			'마이너스 재고는 의미가 없다.
			if FRectIncludePreOrderNo = "Y" then
				sqlStr = sqlStr + " having (((case when IsNULL(s.realstock,0) < 0 then 0 else IsNULL(s.realstock,0) end) - sum(Case when m.statecd in ('0', '1', '6') then d.baljuitemno else 0 end ) + IsNULL(s.preorderno,0) + IsNull(sum(distinct  Case when  ujm.baljucode is not null and ujm.statecd = '9' then ujd.baljuitemno else 0 end ),0)) < 0) "
			else
				sqlStr = sqlStr + " having (((case when IsNULL(s.realstock,0) < 0 then 0 else IsNULL(s.realstock,0) end) - sum(Case when m.statecd in ('0', '1', '6') then d.baljuitemno else 0 end )) < 0) "
			end if
		end if

		sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid"
		If FGroupByBaljuCode = True Then
			sqlStr = sqlStr + ",m.baljucode"
		End If
		''response.write sqlStr & "<br><br>"

		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljusumItem
				FItemList(i).FMakerid          	= rsget("makerid")
				FItemList(i).FItemGubun        	= rsget("itemgubun")
				FItemList(i).FItemId           	= rsget("itemid")
				FItemList(i).FItemoption       	= rsget("itemoption")
				FItemList(i).FJupsuCount       	= rsget("jupsucnt")
				FItemList(i).FCount            	= rsget("cnt")
				FItemList(i).FItemName         	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionname   	= db2html(rsget("itemoptionname"))
                FItemList(i).Frealstock        	= rsget("realstock")

                FItemList(i).Fmwdiv				= rsget("mwdiv")
                FItemList(i).Fpreorderno        = rsget("preorderno") + (rsget("preorderno9") - rsget("preordernofix9"))
                FItemList(i).Fpreordernofix     = rsget("preordernofix")
                FItemList(i).FreipgoMayDate     = rsget("reipgoMayDate")

                FItemList(i).Fshopid				= rsget("shopid")
                FItemList(i).Fshopname				= db2html(rsget("shopname"))
                FItemList(i).Fbaljucode				= rsget("baljucode")
				FItemList(i).Fbaljucodecnt			= rsget("baljucodecnt")

                FItemList(i).Fupcheorderlinkcode	= rsget("upcheorderlinkcode")

                FItemList(i).Fdeliverytype	= rsget("deliverytype")
                FItemList(i).Fcentermwdiv	= rsget("centermwdiv")

				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("offimgsmall")

				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""

				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall

				FItemList(i).FpriceCnt = rsget("priceCnt")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public sub GetFranBalju2UpcheBaljuSheetList
		dim i,sqlStr
		sqlStr = "select top " + CStr(FPageSize) + " m.idx, m.baljucode, m.regdate, m.scheduledate, m.ipgodate "
		sqlStr = sqlStr + " ,m.baljuid, m.baljuname, m.targetid, targetname"
		sqlStr = sqlStr + " ,m.reguser, m.regname, m.divcode, m.statecd"
		sqlStr = sqlStr + " ,sum(d.sellcash*d.realitemno) as totalsellcash, sum(d.buycash*d.realitemno) as totalbuycash"
		sqlStr = sqlStr + " ,sum(d.sellcash*d.baljuitemno) as jumunsellcash, sum(d.buycash*d.baljuitemno) as jumunbuycash"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where m.idx<>0"

		if FRectStatecd<>"" then
			sqlStr = sqlStr + " and m.statecd='" + FRectStatecd + "'"
		end if
		sqlStr = sqlStr + " and m.idx=d.masteridx"

		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and m.targetid='" + FRectBaljuId + "'"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		if FRectDivCode<>"" then
			sqlStr = sqlStr + " and m.divcode='" + FRectDivCode + "'"
		end if

		if FRectDivCodeUnder<>"" then
			sqlStr = sqlStr + " and m.divcode<'" + FRectDivCodeUnder + "'"
		end if

		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " group by m.idx, m.baljucode, m.regdate, m.scheduledate, m.ipgodate"
		sqlStr = sqlStr + " ,m.baljuid, m.baljuname, m.targetid, targetname"
		sqlStr = sqlStr + " ,m.reguser, m.regname, m.divcode, statecd"
		sqlStr = sqlStr + " order by m.idx desc"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem
				FItemList(i).Fidx    = rsget("idx")
				FItemList(i).Fbaljucode  = rsget("baljucode")
				FItemList(i).FRegdate  = rsget("regdate")
				FItemList(i).Fscheduledate = rsget("scheduledate")
				FItemList(i).Fipgodate = rsget("ipgodate")
				FItemList(i).Fbaljuid = rsget("baljuid")
				FItemList(i).Ftargetid = rsget("targetid")
				FItemList(i).Fbaljuname = rsget("baljuname")
				FItemList(i).Ftargetname = rsget("targetname")
				FItemList(i).Freguser = rsget("reguser")
				FItemList(i).Fregname = rsget("regname")
				FItemList(i).Fdivcode = rsget("divcode")
				FItemList(i).Ftotalsellcash = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash = rsget("totalbuycash")
				FItemList(i).Fjumunsellcash = rsget("jumunsellcash")
				FItemList(i).Fjumunbuycash = rsget("jumunbuycash")
				FItemList(i).Fstatecd = rsget("statecd")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	'샆발주대비출고금액
	public sub GetFranBaljuVSChulgo
		dim i,sqlStr

        sqlStr = " select m.baljuid, m.baljuname, d.itemgubun "
        sqlStr = sqlStr + " ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.sellcash*d.baljuitemno else 0 end) as jumunsellcash "
        sqlStr = sqlStr + " ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.suplycash*d.baljuitemno else 0 end) as jumunsuplycash "
        sqlStr = sqlStr + " ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.buycash*d.baljuitemno else 0 end) as jumunbuycash "
        sqlStr = sqlStr + " ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.baljuitemno else 0 end) as jumunitemno "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.sellcash*d.realitemno else 0 end) as totalsellcash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.suplycash*d.realitemno else 0 end) as totalsuplycash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.buycash*d.realitemno else 0 end) as totalbuycash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.realitemno else 0 end) as totalitemno "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno < 0) then d.sellcash*d.realitemno else 0 end) as totalreturnsellcash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno < 0) then d.suplycash*d.realitemno else 0 end) as totalreturnsuplycash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno < 0) then d.buycash*d.realitemno else 0 end) as totalreturnbuycash "
        sqlStr = sqlStr + " ,sum(case when (m.statecd = '7' and d.baljuitemno < 0) then d.realitemno else 0 end) as totalreturnitemno "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx=d.masteridx "
        sqlStr = sqlStr + " and m.deldt is null "
        sqlStr = sqlStr + " and d.deldt is null "
        sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + " and m.ipgodate is not null "

		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and m.baljuid = '" + CStr(FRectBaljuId) + "' "
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and m.ipgodate>='" + CStr(FRectStartDate) + "' "
		end if

		if FRectEndDate<>"" then
			sqlStr = sqlStr + " and m.ipgodate<'" + CStr(FRectEndDate) + "' "
		end if

        sqlStr = sqlStr + " group by m.baljuid, m.baljuname, d.itemgubun "
        sqlStr = sqlStr + " order by m.baljuid, m.baljuname, d.itemgubun "
        'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem
				FItemList(i).Fbaljuid = rsget("baljuid")
				FItemList(i).Fbaljuname = rsget("baljuname")
				FItemList(i).Fitemgubun = rsget("itemgubun")
				FItemList(i).Fjumunsellcash = rsget("jumunsellcash")
				FItemList(i).Fjumunsuplycash = rsget("jumunsuplycash")
				FItemList(i).Fjumunbuycash = rsget("jumunbuycash")
				FItemList(i).Fjumunitemno = rsget("jumunitemno")
				FItemList(i).Ftotalsellcash = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
				FItemList(i).Ftotalbuycash = rsget("totalbuycash")
				FItemList(i).Ftotalitemno = rsget("totalitemno")
				FItemList(i).Ftotalreturnsellcash = rsget("totalreturnsellcash")
				FItemList(i).Ftotalreturnsuplycash = rsget("totalreturnsuplycash")
				FItemList(i).Ftotalreturnbuycash = rsget("totalreturnbuycash")
				FItemList(i).Ftotalreturnitemno = rsget("totalreturnitemno")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

        '샆발주대비출고금액(상품별)
	public sub GetFranBaljuVSChulgoByItem
		dim i,sqlStr

        sqlStr = " select top 10 T.baljuid, T.baljuname, T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.jumunbuycash, T.totalbuycash, T.jumunsuplycash, T.totalsuplycash "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select m.baljuid, m.baljuname, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname "
        sqlStr = sqlStr + "         ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.sellcash*d.baljuitemno else 0 end) as jumunsellcash "
        sqlStr = sqlStr + "         ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.suplycash*d.baljuitemno else 0 end) as jumunsuplycash "
        sqlStr = sqlStr + "         ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.buycash*d.baljuitemno else 0 end) as jumunbuycash "
        sqlStr = sqlStr + "         ,sum(case when (left(m.baljucode,2) <> 'RJ' and d.baljuitemno > 0) then d.baljuitemno else 0 end) as jumunitemno "
        sqlStr = sqlStr + "         ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.sellcash*d.realitemno else 0 end) as totalsellcash "
        sqlStr = sqlStr + "         ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.suplycash*d.realitemno else 0 end) as totalsuplycash "
        sqlStr = sqlStr + "         ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.buycash*d.realitemno else 0 end) as totalbuycash "
        sqlStr = sqlStr + "         ,sum(case when (m.statecd = '7' and d.baljuitemno > 0) then d.realitemno else 0 end) as totalitemno "
        sqlStr = sqlStr + "     from [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + "     where 1 = 1 "
        sqlStr = sqlStr + "     and m.idx=d.masteridx "
        sqlStr = sqlStr + "     and m.deldt is null "
        sqlStr = sqlStr + "     and d.deldt is null "
        sqlStr = sqlStr + "     and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + "     and m.ipgodate is not null "

		if FRectBaljuId<>"" then
			sqlStr = sqlStr + "     and m.baljuid = '" + CStr(FRectBaljuId) + "' "
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + "     and m.ipgodate>='" + CStr(FRectStartDate) + "' "
		end if

		if FRectEndDate<>"" then
			sqlStr = sqlStr + "     and m.ipgodate<'" + CStr(FRectEndDate) + "' "
		end if

        sqlStr = sqlStr + "     group by m.baljuid, m.baljuname, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " order by (T.jumunbuycash - T.totalbuycash) desc, T.baljuid, T.baljuname, T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname "

        'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem
				FItemList(i).Fbaljuid = rsget("baljuid")
				FItemList(i).Fbaljuname = rsget("baljuname")
				FItemList(i).FItemGubun         = rsget("itemgubun")
				FItemList(i).FItemId            = rsget("itemid")
				FItemList(i).FItemoption        = rsget("itemoption")
				FItemList(i).FItemName          = db2html(rsget("itemname"))
				FItemList(i).FItemOptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Ftotalbuycash     = rsget("totalbuycash")
				FItemList(i).Fjumunbuycash     = rsget("jumunbuycash")
				FItemList(i).Ftotalsuplycash     = rsget("totalsuplycash")
				FItemList(i).Fjumunsuplycash     = rsget("jumunsuplycash")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetOrderSheetBaljulist()
		dim i,sqlStr
		sqlStr = " select c.prtidx, d.makerid, d.itemgubun, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.sellcash, sum(d.realitemno) as realitemno,"
		sqlStr = sqlStr + " i.sellyn,i.dispyn,i.limityn,i.limitno,i.limitsold,i.smallimage,s.offimgsmall "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " on d.makerid=c.userid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " on d.itemgubun=s.itemgubun and d.itemid=s.shopitemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " where d.masteridx in (" + CStr(FRectIdxArr) + ")"
		sqlStr = sqlStr + " and d.deldt is Null"
		sqlStr = sqlStr + " group by c.prtidx, d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.sellcash, i.sellyn,i.dispyn,i.limityn,i.limitno,i.limitsold,i.smallimage,s.offimgsmall"
		sqlStr = sqlStr + " order by c.prtidx, d.makerid, d.itemgubun, d.itemid, d.itemoption"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem
				FItemList(i).FRectIsFixed    = true
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Frealitemno    = rsget("realitemno")
				FItemList(i).FMakerid	= rsget("makerid")
				FItemList(i).Fprtidx	= rsget("prtidx")
				FItemList(i).FonlineSellyn = rsget("sellyn")
				FItemList(i).FonlineDispyn = rsget("dispyn")
				FItemList(i).FonlineLimityn = rsget("limityn")
				FItemList(i).FonlineLimitno = rsget("limitno")
				FItemList(i).FonlineLimitsold = rsget("limitsold")
				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("smallimage")
				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""
				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall


				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//admin/fran/jumuninputedit.asp
	public Sub GetOrderSheetDetail()
		dim i,sqlStr

		sqlStr = " select"
		sqlStr = sqlStr & " d.idx,d.masteridx,d.itemgubun,d.makerid,d.itemid,d.itemoption,d.itemname,d.itemoptionname"
		sqlStr = sqlStr & " ,d.sellcash,d.suplycash,isnull(d.buycash,0) as buycash,isnull(d.baljuitemno,0) as baljuitemno,isnull(d.realitemno,0) as realitemno"
		sqlStr = sqlStr & " ,d.regdate,d.updt,d.deldt"
		sqlStr = sqlStr & " ,d.baljudiv,d.comment,d.ipgoflag,d.defaultmaginflag,d.buymaginflag,d.suplymaginflag"
		sqlStr = sqlStr & " ,d.packingstate,d.boxsongjangno,d.upcheorderlinkcode,d.realbaljuitemno,d.checkitemno"
		sqlStr = sqlStr & " ,d.foreign_sellcash,d.foreign_suplycash"
		sqlStr = sqlStr + " , i.mwdiv, i.smallimage, i.deliverytype "
		sqlStr = sqlStr + " , IsNULL(i.sellcash,0) as onlinesellcash, IsNULL(i.buycash,0) as onlinebuycash"
		sqlStr = sqlStr + " , IsNULL(sd.defaultmargin,0) as shopdefaultmargin, IsNULL(sd.defaultsuplymargin,0) as shopdefaultsuplymargin "
		sqlStr = sqlStr + " , IsNULL(sd.chargediv,'0') as offchargediv"
		sqlStr = sqlStr + " , s.barcode, si.centermwdiv, si.offimgmain, si.offimglist, si.offimgsmall"
		sqlStr = sqlStr + " , dl.detail_status,dl.detail_description, isNull(s.upchemanagecode,'') AS upchemanagecode, IsNull(si.orgsellprice,0) as orgsellprice "
		sqlStr = sqlStr + " , i.basicimage, si.offimgmain "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
	    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item si"
	    sqlStr = sqlStr + "     on d.itemgubun<>'10' and d.itemgubun=si.itemgubun and d.itemid=si.shopitemid and d.itemoption=si.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on d.itemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sd "
		sqlStr = sqlStr + "     on sd.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + "     and d.makerid=sd.makerid"
		sqlStr = sqlStr + " left join db_storage.dbo.tbl_ordersheet_detail_log dl "
		sqlStr = sqlStr + " 	on d.idx = dl.detail_idx "

		sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and d.deldt is Null"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if

		sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption"		' 다른매뉴와 동일하게 맞춤

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).fforeign_sellcash            = rsget("foreign_sellcash")
				FItemList(i).fforeign_suplycash      = rsget("foreign_suplycash")
				FItemList(i).FRectIsFixed    = FRectIsFixed
				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Fsuplycash      = rsget("suplycash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fupdt         = rsget("updt")
				FItemList(i).Fdeldt        = rsget("deldt")
				FItemList(i).Fbaljudiv     = rsget("baljudiv")
				FItemList(i).Fcomment      = db2html(rsget("comment"))
				FItemList(i).FMakerid	= rsget("makerid")
				FItemList(i).FItemDefaultMwDiv	= rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Fonlinesellcash = rsget("onlinesellcash")
				FItemList(i).Fonlinebuycash  = rsget("onlinebuycash")
				FItemList(i).Fshopdefaultmargin = rsget("shopdefaultmargin")
				FItemList(i).Fshopdefaultsuplymargin = rsget("shopdefaultsuplymargin")
				FItemList(i).FoffChargeDiv = rsget("offchargediv")
				FItemList(i).Fipgoflag = rsget("ipgoflag")
				FItemList(i).Fdefaultmaginflag = rsget("defaultmaginflag")
				FItemList(i).Fbuymaginflag = rsget("buymaginflag")
				FItemList(i).Fsuplymaginflag = rsget("suplymaginflag")
				FItemList(i).FPublicBarcode = rsget("barcode")

				FItemList(i).Fsmallimage = rsget("smallimage")
				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage

				FItemList(i).Fbasicimage = rsget("basicimage")
				if isnull(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				if FItemList(i).Fbasicimage<>"" then FItemList(i).Fbasicimage     = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fbasicimage

				FItemList(i).FOffimgMain	= rsget("offimgmain")
					if isnull(FItemList(i).FOffimgMain) then FItemList(i).FOffimgMain=""
				FItemList(i).FOffimgList	= rsget("offimglist")
					if isnull(FItemList(i).FOffimgList) then FItemList(i).FOffimgList=""
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if isnull(FItemList(i).FOffimgSmall) then FItemList(i).FOffimgSmall=""

				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgSmall

				FItemList(i).Fdetail_status = rsget("detail_status")
				FItemList(i).Fdetail_description = rsget("detail_description")
				FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
				FItemList(i).Fboxsongjangno  = rsget("boxsongjangno")
				FItemList(i).FUpcheManageCode = rsget("upchemanagecode")
				FItemList(i).Forgsellprice = rsget("orgsellprice")

				if isnull(FItemList(i).Fcheckitemno) then FItemList(i).Fcheckitemno = 0

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	' 상품별주문리스트		' 2020.09.14 한용민 생성
	'/admin/newstorage/itemOrderList.asp
	' 밑에 함수를 수정할경우 GetItemOrderListNotPaging 함수도 똑같이 수정해야 한다.
	public Sub GetItemOrderList()
		dim sqlStr,i, AddSql

		AddSql=""
		if FRectdatetype="regdate" or FRectdatetype="scheduledate" then
			if FRectStartDate<>"" and FRectEndDate<>"" then
				if FRectStartDate<>"" then
					AddSql = AddSql & " and m."& FRectdatetype &">='"& FRectStartDate &"'"
				end if
				if FRectEndDate<>"" then
					AddSql = AddSql & " and m."& FRectdatetype &"<'"& FRectEndDate &"'"
				end if
			end if
		end if
		if FRectbaljucode<>"" then
			AddSql = AddSql & " and m.baljucode='"& FRectbaljucode &"'"
		end if
		if FRectblinkcode<>"" then
			AddSql = AddSql & " and m.blinkcode='"& FRectblinkcode &"'"
		end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
		if FRectmakerid<>"" then
			AddSql = AddSql & " and d.makerid='"& FRectmakerid &"'"
		end if
		if FRectmwdiv<>"" then
			AddSql = AddSql & " and i.mwdiv='"& FRectmwdiv &"'"
		end if
		if (FRectBrandPurchaseType <> "") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				AddSql = AddSql + " 	and pp.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				AddSql = AddSql & " 	and pp.purchasetype in ('3','5','6')"
			else
				AddSql = AddSql + " 	and pp.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if
		if FRectStatecd<>"" then
			if FRectStatecd="foreign0" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=0"
			elseif FRectStatecd="foreign3" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=3"
			elseif FRectStatecd="foreign7" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=7"
			elseif FRectStatecd=" " then
				AddSql = AddSql + " and m.statecd='" + FRectStatecd + "' and m.foreign_statecd is null"
			elseif FRectStatecd="preorder" then		'//기주문상태
				AddSql = AddSql + " and m.statecd<9 and m.scheduledate >= dateadd(m,-2,getdate())"
			elseif FRectStatecd="before1" then
				AddSql = AddSql + " and m.statecd < '1' "
			else
				AddSql = AddSql + " and m.statecd='" + FRectStatecd + "'"
			end if
		end if
		if (FRecttplgubun <> "") then
			if (FRecttplgubun = "3X") then
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FRecttplgubun) + "' "
			end if
		end if
		if (FRectproductidx <> "") then
			AddSql = AddSql & " and m.idx in ("
			AddSql = AddSql & " 	select"
			AddSql = AddSql & " 	pl.linkIdx"
			AddSql = AddSql & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
			AddSql = AddSql & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
			AddSql = AddSql & " 		on pm.idx=pl.ppMasterIdx"
			AddSql = AddSql & " 	where pm.deldt is null"
			AddSql = AddSql & " 	and pl.deldt is null"
			AddSql = AddSql & " 	and pm.idx="& FRectproductidx &""
			AddSql = AddSql & " )"
		end if

		sqlStr = " select count(m.idx) as cnt, CEILING(CAST(Count(m.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master as m with (nolock)"
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx"
		sqlStr = sqlStr & " 	and d.deldt is Null"
		sqlStr = sqlStr & " left Join db_partner.dbo.tbl_partner p with (nolock)"
		sqlStr = sqlStr & " 	on m.targetid=p.id"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr & " 	on d.itemgubun='10' and d.itemid=i.itemid" 
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.makerid=pp.id"
		sqlStr = sqlStr & " where m.deldt is null"
		sqlStr = sqlStr & " and m.divcode in ('301','302') " & AddSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top "&FPageSize*FCurrPage
		sqlStr = sqlStr & " m.baljucode"
		sqlStr = sqlStr & " , isnull(("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	string_agg(ppMasterIdx,',') WITHIN GROUP(ORDER BY ppMasterIdx asc) as ppmasteridx"
		sqlStr = sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		sqlStr = sqlStr & " 		on pm.idx=pl.ppMasterIdx"
		sqlStr = sqlStr & " 	where pm.deldt is null"
		sqlStr = sqlStr & " 	and pl.deldt is null"
		sqlStr = sqlStr & " 	and m.idx=pl.linkIdx"
		sqlStr = sqlStr & " 	group by linkIdx"
		'sqlStr = sqlStr & " 	select top 1"
		'sqlStr = sqlStr & " 	pl.ppmasteridx"
		'sqlStr = sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		'sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		'sqlStr = sqlStr & " 		on pm.idx=pl.ppMasterIdx"
		'sqlStr = sqlStr & " 	where pm.deldt is null"
		'sqlStr = sqlStr & " 	and pl.deldt is null"
		'sqlStr = sqlStr & " 	and m.idx=pl.linkIdx"
		'sqlStr = sqlStr & " 	order by pl.ppmasteridx desc"
		sqlStr = sqlStr & " ),'') as productidxArr"
		sqlStr = sqlStr & " , isnull(ep.reportidx,'') as reportidx, m.regdate, m.scheduledate"
		sqlStr = sqlStr & " , (case"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='0' then '업체접수(견적요청)'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='3' then '업체접수확인'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='7' then '업체접수완료'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and (m.foreign_statecd is null or m.foreign_statecd='') then '주문서작성중'"
		sqlStr = sqlStr & " 	when m.statecd='0' then '주문접수'"
		sqlStr = sqlStr & " 	when m.statecd='1' then '주문확인'"
		sqlStr = sqlStr & " 	when m.statecd='2' then '입금대기'"
		sqlStr = sqlStr & " 	when m.statecd='5' then '배송준비'"
		sqlStr = sqlStr & " 	when m.statecd='6' then '출고대기'"
		sqlStr = sqlStr & " 	when m.statecd='7' then '출고완료'"
		sqlStr = sqlStr & " 	when m.statecd='8' then '검품완료(입고대기)'"
		sqlStr = sqlStr & " 	when m.statecd='9' then '입고완료' else '' end) as statecdname"
		sqlStr = sqlStr & " , isnull(m.blinkcode,'') as blinkcode"
		sqlStr = sqlStr & " , d.makerid, d.itemgubun, d.itemid, d.itemoption"
		sqlStr = sqlStr & " , (db_item.[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)) as tenbarcode"
		sqlStr = sqlStr & " , isnull(s.barcode,'') as barcode"
		sqlStr = sqlStr & " , replace(replace(replace(replace(replace(s.upchemanagecode,char(9),''),char(10),''),char(13),''),'""',''),'''','') as upchemanagecode"
		sqlStr = sqlStr & " , replace(replace(replace(replace(replace(d.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','') as itemname"
		sqlStr = sqlStr & " , isnull(replace(replace(replace(replace(replace(d.itemoptionname,char(9),''),char(10),''),char(13),''),'""',''),'''',''),'') as itemoptionname"
		sqlStr = sqlStr & " , isnull(d.sellcash,0) as sellcash, isnull(d.buycash,0) as buycash, i.mwdiv, isnull(d.baljuitemno,0) as baljuitemno"
		sqlStr = sqlStr & " , isnull(d.realitemno,0) as realitemno, isnull(d.checkitemno,0) as checkitemno"
		sqlStr = sqlStr & " , pc.pcomm_name as purchaseTypename"
		sqlStr = sqlStr & " , isNULL(cl.cateName,'미지정') as cateName"
		sqlStr = sqlStr & " , isnull(ls.lastIpgoDate,'') as lastIpgoDate"
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master as m with (nolock)"
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx"
		sqlStr = sqlStr & " 	and d.deldt is Null"
		sqlStr = sqlStr & " left Join db_partner.dbo.tbl_partner p with (nolock)"
		sqlStr = sqlStr & " 	on m.targetid=p.id"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr & " 	on d.itemgubun='10' and d.itemid=i.itemid" 
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)"
		sqlStr = sqlStr & " 	on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock)"
		sqlStr = sqlStr & " 	on m.idx = ep.scmlinkNo and ep.isUsing =1"
		sqlStr = sqlStr & " 	and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69)"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_display_cate_item as ci with (nolock)"
		sqlStr = sqlStr & " 	ON d.itemid = ci.itemid AND ci.isDefault='y'"
		sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_display_cate] as cl with (nolock)"
		sqlStr = sqlStr & " 	ON Left(ci.catecode,3)=cl.catecode"
		sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] ls with (nolock)"		' 월별누적재고
		sqlStr = sqlStr & " 	on ls.yyyymm=convert(varchar(7),getdate(),121)"
		sqlStr = sqlStr & " 	and d.itemgubun=ls.itemgubun"
		sqlStr = sqlStr & " 	and d.itemid=ls.itemid"
		sqlStr = sqlStr & " 	and d.itemoption=ls.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.makerid=pp.id"
	    sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
	    sqlStr = sqlStr + " 	on pc.pcomm_isusing='Y' and pc.pcomm_group='purchasetype'"
	    sqlStr = sqlStr + " 	and p.purchaseType=pc.pcomm_cd"
		sqlStr = sqlStr & " where m.deldt is null"
		sqlStr = sqlStr & " and m.divcode in ('301','302') " & AddSql
		sqlStr = sqlStr & " order by m.baljucode desc, d.makerid asc, d.itemgubun asc, d.itemid asc, d.itemoption asc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

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
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).fbaljucode      = rsget("baljucode")
				FItemList(i).fproductidxArr      = rsget("productidxArr")
				FItemList(i).freportidx      = rsget("reportidx")
				FItemList(i).fregdate      = rsget("regdate")
				FItemList(i).fscheduledate      = rsget("scheduledate")
				FItemList(i).fstatecdname      = rsget("statecdname")
				FItemList(i).fblinkcode      = rsget("blinkcode")
				FItemList(i).fmakerid      = rsget("makerid")
				FItemList(i).fitemgubun      = rsget("itemgubun")
				FItemList(i).fitemid      = rsget("itemid")
				FItemList(i).fitemoption      = rsget("itemoption")
				FItemList(i).ftenbarcode      = rsget("tenbarcode")
				FItemList(i).fbarcode      = rsget("barcode")
				FItemList(i).fupchemanagecode      = db2html(rsget("upchemanagecode"))
				FItemList(i).fitemname      = db2html(rsget("itemname"))
				FItemList(i).fitemoptionname      = db2html(rsget("itemoptionname"))
				FItemList(i).fsellcash      = rsget("sellcash")
				FItemList(i).fbuycash      = rsget("buycash")
				FItemList(i).fmwdiv      = rsget("mwdiv")
				FItemList(i).fbaljuitemno      = rsget("baljuitemno")
				FItemList(i).frealitemno      = rsget("realitemno")
				FItemList(i).fcheckitemno      = rsget("checkitemno")
				FItemList(i).fpurchaseTypename      = rsget("purchaseTypename")
				FItemList(i).fcateName      = db2html(rsget("cateName"))
				FItemList(i).flastIpgoDate      = rsget("lastIpgoDate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	' 상품별주문리스트		' 2020.09.14 한용민 생성
	'/admin/newstorage/itemOrderList_excel.asp
	' 밑에 함수를 수정할경우 GetItemOrderList 함수도 똑같이 수정해야 한다.
	public Sub GetItemOrderListNotPaging()
		dim sqlStr,i, AddSql

		AddSql=""
		if FRectdatetype="regdate" or FRectdatetype="scheduledate" then
			if FRectStartDate<>"" and FRectEndDate<>"" then
				if FRectStartDate<>"" then
					AddSql = AddSql & " and m."& FRectdatetype &">='"& FRectStartDate &"'"
				end if
				if FRectEndDate<>"" then
					AddSql = AddSql & " and m."& FRectdatetype &"<'"& FRectEndDate &"'"
				end if
			end if
		end if
		if FRectbaljucode<>"" then
			AddSql = AddSql & " and m.baljucode='"& FRectbaljucode &"'"
		end if
		if FRectblinkcode<>"" then
			AddSql = AddSql & " and m.blinkcode='"& FRectblinkcode &"'"
		end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
		if FRectmakerid<>"" then
			AddSql = AddSql & " and d.makerid='"& FRectmakerid &"'"
		end if
		if FRectmwdiv<>"" then
			AddSql = AddSql & " and i.mwdiv='"& FRectmwdiv &"'"
		end if
		if (FRectBrandPurchaseType <> "") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				AddSql = AddSql + " 	and pp.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				AddSql = AddSql & " 	and pp.purchasetype in ('3','5','6')"
			else
				AddSql = AddSql + " 	and pp.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if
		if FRectStatecd<>"" then
			if FRectStatecd="foreign0" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=0"
			elseif FRectStatecd="foreign3" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=3"
			elseif FRectStatecd="foreign7" then
				AddSql = AddSql + " and m.statecd=' ' and m.foreign_statecd=7"
			elseif FRectStatecd=" " then
				AddSql = AddSql + " and m.statecd='" + FRectStatecd + "' and m.foreign_statecd is null"
			elseif FRectStatecd="preorder" then		'//기주문상태
				AddSql = AddSql + " and m.statecd<9 and m.scheduledate >= dateadd(m,-2,getdate())"
			elseif FRectStatecd="before1" then
				AddSql = AddSql + " and m.statecd < '1' "
			else
				AddSql = AddSql + " and m.statecd='" + FRectStatecd + "'"
			end if
		end if
		if (FRecttplgubun <> "") then
			if (FRecttplgubun = "3X") then
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FRecttplgubun) + "' "
			end if
		end if
		if (FRectproductidx <> "") then
			AddSql = AddSql & " and m.idx in ("
			AddSql = AddSql & " 	select"
			AddSql = AddSql & " 	pl.linkIdx"
			AddSql = AddSql & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
			AddSql = AddSql & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
			AddSql = AddSql & " 		on pm.idx=pl.ppMasterIdx"
			AddSql = AddSql & " 	where pm.deldt is null"
			AddSql = AddSql & " 	and pl.deldt is null"
			AddSql = AddSql & " 	and pm.idx="& FRectproductidx &""
			AddSql = AddSql & " )"
		end if

		sqlStr = " select top "&FPageSize*FCurrPage
		sqlStr = sqlStr & " m.baljucode"
		sqlStr = sqlStr & " , isnull(("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	string_agg(ppMasterIdx,',') WITHIN GROUP(ORDER BY ppMasterIdx asc) as ppmasteridx"
		sqlStr = sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		sqlStr = sqlStr & " 		on pm.idx=pl.ppMasterIdx"
		sqlStr = sqlStr & " 	where pm.deldt is null"
		sqlStr = sqlStr & " 	and pl.deldt is null"
		sqlStr = sqlStr & " 	and m.idx=pl.linkIdx"
		sqlStr = sqlStr & " 	group by linkIdx"
		'sqlStr = sqlStr & " 	select top 1"
		'sqlStr = sqlStr & " 	pl.ppmasteridx"
		'sqlStr = sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		'sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		'sqlStr = sqlStr & " 		on pm.idx=pl.ppMasterIdx"
		'sqlStr = sqlStr & " 	where pm.deldt is null"
		'sqlStr = sqlStr & " 	and pl.deldt is null"
		'sqlStr = sqlStr & " 	and m.idx=pl.linkIdx"
		'sqlStr = sqlStr & " 	order by pl.ppmasteridx desc"
		sqlStr = sqlStr & " ),'') as productidxArr"
		sqlStr = sqlStr & " , isnull(ep.reportidx,'') as reportidx, m.regdate, m.scheduledate"
		sqlStr = sqlStr & " , (case"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='0' then '업체접수(견적요청)'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='3' then '업체접수확인'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and m.foreign_statecd='7' then '업체접수완료'"
		sqlStr = sqlStr & " 	when m.statecd=' ' and (m.foreign_statecd is null or m.foreign_statecd='') then '주문서작성중'"
		sqlStr = sqlStr & " 	when m.statecd='0' then '주문접수'"
		sqlStr = sqlStr & " 	when m.statecd='1' then '주문확인'"
		sqlStr = sqlStr & " 	when m.statecd='2' then '입금대기'"
		sqlStr = sqlStr & " 	when m.statecd='5' then '배송준비'"
		sqlStr = sqlStr & " 	when m.statecd='6' then '출고대기'"
		sqlStr = sqlStr & " 	when m.statecd='7' then '출고완료'"
		sqlStr = sqlStr & " 	when m.statecd='8' then '검품완료(입고대기)'"
		sqlStr = sqlStr & " 	when m.statecd='9' then '입고완료' else '' end) as statecdname"
		sqlStr = sqlStr & " , isnull(m.blinkcode,'') as blinkcode"
		sqlStr = sqlStr & " , d.makerid, d.itemgubun, d.itemid, d.itemoption"
		sqlStr = sqlStr & " , (db_item.[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)) as tenbarcode"
		sqlStr = sqlStr & " , isnull(s.barcode,'') as barcode"
		sqlStr = sqlStr & " , replace(replace(replace(replace(replace(s.upchemanagecode,char(9),''),char(10),''),char(13),''),'""',''),'''','') as upchemanagecode"
		sqlStr = sqlStr & " , replace(replace(replace(replace(replace(d.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','') as itemname"
		sqlStr = sqlStr & " , isnull(replace(replace(replace(replace(replace(d.itemoptionname,char(9),''),char(10),''),char(13),''),'""',''),'''',''),'') as itemoptionname"
		sqlStr = sqlStr & " , isnull(d.sellcash,0) as sellcash, isnull(d.buycash,0) as buycash, i.mwdiv, isnull(d.baljuitemno,0) as baljuitemno"
		sqlStr = sqlStr & " , isnull(d.realitemno,0) as realitemno, isnull(d.checkitemno,0) as checkitemno"
		sqlStr = sqlStr & " , pc.pcomm_name as purchaseTypename"
		sqlStr = sqlStr & " , isNULL(cl.cateName,'미지정') as cateName"
		sqlStr = sqlStr & " , isnull(ls.lastIpgoDate,'') as lastIpgoDate"
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master as m with (nolock)"
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx"
		sqlStr = sqlStr & " 	and d.deldt is Null"
		sqlStr = sqlStr & " left Join db_partner.dbo.tbl_partner p with (nolock)"
		sqlStr = sqlStr & " 	on m.targetid=p.id"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr & " 	on d.itemgubun='10' and d.itemid=i.itemid" 
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)"
		sqlStr = sqlStr & " 	on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock)"
		sqlStr = sqlStr & " 	on m.idx = ep.scmlinkNo and ep.isUsing =1"
		sqlStr = sqlStr & " 	and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69)"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_display_cate_item as ci with (nolock)"
		sqlStr = sqlStr & " 	ON d.itemid = ci.itemid AND ci.isDefault='y'"
		sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_display_cate] as cl with (nolock)"
		sqlStr = sqlStr & " 	ON Left(ci.catecode,3)=cl.catecode"
		sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] ls with (nolock)"		' 월별누적재고
		sqlStr = sqlStr & " 	on ls.yyyymm=convert(varchar(7),getdate(),121)"
		sqlStr = sqlStr & " 	and d.itemgubun=ls.itemgubun"
		sqlStr = sqlStr & " 	and d.itemid=ls.itemid"
		sqlStr = sqlStr & " 	and d.itemoption=ls.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.makerid=pp.id"
	    sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
	    sqlStr = sqlStr + " 	on pc.pcomm_isusing='Y' and pc.pcomm_group='purchasetype'"
	    sqlStr = sqlStr + " 	and p.purchaseType=pc.pcomm_cd"
		sqlStr = sqlStr & " where m.deldt is null"
		sqlStr = sqlStr & " and m.divcode in ('301','302') " & AddSql
		sqlStr = sqlStr & " order by m.baljucode desc, d.makerid asc, d.itemgubun asc, d.itemid asc, d.itemoption asc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close
	end sub

	'주문서관리 수출용 상품리스트. 홀쎄일에서 꼿는 주문이 아니여서 해외가격이 없음. 따로 쿼리함.	'/2017.06.15 한용민 수정
	'//admin/newstorage/ordersheet.asp
	public Sub GetOrderSheetDetail_foreign()
		dim i,sqlStr, sqlsearch

		if FRectIdx = "" then exit Sub

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectMakerid + "'"
		end if
		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and d.masteridx=" + FRectIdx + ""
		end if

		sqlStr = " select"
		sqlStr = sqlStr + " d.*, i.mwdiv, i.smallimage, i.deliverytype "
		sqlStr = sqlStr + " , IsNULL(i.sellcash,0) as onlinesellcash, IsNULL(i.buycash,0) as onlinebuycash"
		sqlStr = sqlStr + " , IsNULL(sd.defaultmargin,0) as shopdefaultmargin, IsNULL(sd.defaultsuplymargin,0) as shopdefaultsuplymargin "
		sqlStr = sqlStr + " , IsNULL(sd.chargediv,'0') as offchargediv"
		sqlStr = sqlStr + " , s.barcode, si.centermwdiv, si.offimgsmall"
		sqlStr = sqlStr + " , dl.detail_status,dl.detail_description, isNull(s.upchemanagecode,'') AS upchemanagecode "
		sqlStr = sqlStr & " , isnull(isnull(r.orgprice,f.lcprice),0) as lcprice"
	    sqlStr = sqlStr & " , isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
		sqlStr = sqlStr & " , isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_master m"
		sqlStr = sqlStr & " 	on m.idx = d.masteridx"
	    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item si"
	    sqlStr = sqlStr + "     on d.itemgubun<>'10' and d.itemgubun=si.itemgubun and d.itemid=si.shopitemid and d.itemoption=si.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on d.itemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sd "
		sqlStr = sqlStr + "     on sd.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + "     and d.makerid=sd.makerid"
		sqlStr = sqlStr + " left join db_storage.dbo.tbl_ordersheet_detail_log dl "
		sqlStr = sqlStr + " 	on d.idx = dl.detail_idx "
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_locale_item f " + vbcrlf
		sqlStr = sqlStr & " 	on f.shopid = m.baljuid"
		sqlStr = sqlStr & " 	and d.itemgubun=f.itemgubun " + vbcrlf
		sqlStr = sqlStr & " 	and d.itemid=f.shopitemid " + vbcrlf
		sqlStr = sqlStr & " 	and d.itemoption=f.itemoption " + vbcrlf
	    sqlStr = sqlStr & "  left join db_item.[dbo].[tbl_item_multiLang] Lni"
        sqlStr = sqlStr & "  	on Lni.countryCd='EN'"
        sqlStr = sqlStr & "  	and d.itemgubun='10'"
        sqlStr = sqlStr & "  	and d.itemid=Lni.itemid"
        sqlStr = sqlStr & "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno"
        sqlStr = sqlStr & "  	on Lno.countryCd='EN'"
        sqlStr = sqlStr & "  	and d.itemgubun='10'"
        sqlStr = sqlStr & "  	and d.itemid=Lno.itemid"
        sqlStr = sqlStr & "  	and d.itemoption=Lno.itemoption"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " and d.deldt is Null"
		sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption"

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).flcprice            = rsget("lcprice")
				FItemList(i).fforeign_sellcash            = rsget("foreign_sellcash")
				FItemList(i).fforeign_suplycash      = rsget("foreign_suplycash")
				FItemList(i).FRectIsFixed    = FRectIsFixed
				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Fsuplycash      = rsget("suplycash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fupdt         = rsget("updt")
				FItemList(i).Fdeldt        = rsget("deldt")
				FItemList(i).Fbaljudiv     = rsget("baljudiv")
				FItemList(i).Fcomment      = db2html(rsget("comment"))
				FItemList(i).FMakerid	= rsget("makerid")
				FItemList(i).FItemDefaultMwDiv	= rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Fonlinesellcash = rsget("onlinesellcash")
				FItemList(i).Fonlinebuycash  = rsget("onlinebuycash")
				FItemList(i).Fshopdefaultmargin = rsget("shopdefaultmargin")
				FItemList(i).Fshopdefaultsuplymargin = rsget("shopdefaultsuplymargin")
				FItemList(i).FoffChargeDiv = rsget("offchargediv")
				FItemList(i).Fipgoflag = rsget("ipgoflag")
				FItemList(i).Fdefaultmaginflag = rsget("defaultmaginflag")
				FItemList(i).Fbuymaginflag = rsget("buymaginflag")
				FItemList(i).Fsuplymaginflag = rsget("suplymaginflag")
				FItemList(i).FPublicBarcode = rsget("barcode")
				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("offimgsmall")
				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""
				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall
				FItemList(i).Fdetail_status = rsget("detail_status")
				FItemList(i).Fdetail_description = rsget("detail_description")
				FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
				FItemList(i).Fboxsongjangno  = rsget("boxsongjangno")
				FItemList(i).FUpcheManageCode = rsget("upchemanagecode")

				if isnull(FItemList(i).Fcheckitemno) then FItemList(i).Fcheckitemno = 0

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 영문주문서
	public Sub GetOrderSheetDetail_ENG()
		dim i,sqlStr, sqlsearch

		if FRectIdx = "" then exit Sub

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectMakerid + "'"
		end if
		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and d.masteridx=" + FRectIdx + ""
		end if

		sqlStr = " select"
		sqlStr = sqlStr + " d.*, i.mwdiv, i.smallimage, i.deliverytype "
		sqlStr = sqlStr + " , IsNULL(i.sellcash,0) as onlinesellcash, IsNULL(i.buycash,0) as onlinebuycash"
		sqlStr = sqlStr + " , s.barcode, si.centermwdiv, si.offimgsmall"
		sqlStr = sqlStr + " , isNull(s.upchemanagecode,'') AS upchemanagecode "
		sqlStr = sqlStr & " , isnull(b.buyitemprice,0) as lcprice"
	    sqlStr = sqlStr & " , isnull(b.buyitemname,d.itemname) as lcitemname" + vbcrlf
		sqlStr = sqlStr & " , isnull(b.buyitemoptionname,d.itemoptionname) as lcitemoptionname " + vbcrlf
		sqlStr = sqlStr & " , isnull(b.makerid,d.makerid) as buymakerid " + vbcrlf
		sqlStr = sqlStr & " , isnull(b.currencyUnit,'USD') as buycurrencyUnit " + vbcrlf
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr & " 	on m.idx = d.masteridx"
	    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item si with (nolock)"
	    sqlStr = sqlStr + "     on d.itemgubun<>'10' and d.itemgubun=si.itemgubun and d.itemid=si.shopitemid and d.itemoption=si.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.itemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr & "	left join [db_shop].[dbo].[tbl_buy_item] b with (nolock)"
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and d.itemgubun = b.itemgubun "
		sqlStr = sqlStr & "			and d.itemid = b.buyitemid "
		sqlStr = sqlStr & "			and d.itemoption = b.itemoption "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " and d.deldt is Null"
		sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption"

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).flcprice            = rsget("lcprice")
				FItemList(i).FRectIsFixed    = FRectIsFixed
				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Fsuplycash      = rsget("suplycash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fupdt         = rsget("updt")
				FItemList(i).Fdeldt        = rsget("deldt")
				FItemList(i).Fbaljudiv     = rsget("baljudiv")
				FItemList(i).Fcomment      = db2html(rsget("comment"))
				FItemList(i).FMakerid	= rsget("buymakerid")
				FItemList(i).FItemDefaultMwDiv	= rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Fonlinesellcash = rsget("onlinesellcash")
				FItemList(i).Fonlinebuycash  = rsget("onlinebuycash")
				FItemList(i).Fipgoflag = rsget("ipgoflag")
				FItemList(i).Fdefaultmaginflag = rsget("defaultmaginflag")
				FItemList(i).Fbuymaginflag = rsget("buymaginflag")
				FItemList(i).Fsuplymaginflag = rsget("suplymaginflag")
				FItemList(i).FPublicBarcode = rsget("barcode")
				FItemList(i).Fsmallimage = rsget("smallimage")
				FItemList(i).Foffimgsmall = rsget("offimgsmall")
				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""
				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall
				FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
				FItemList(i).Fboxsongjangno  = rsget("boxsongjangno")
				FItemList(i).FUpcheManageCode = rsget("upchemanagecode")
				FItemList(i).FcurrencyUnit = rsget("buycurrencyUnit")

				if isnull(FItemList(i).Fcheckitemno) then FItemList(i).Fcheckitemno = 0

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	' /partner/storage/orderlist_barcode.asp	' 2017.12.22 한용민 생성
	public Sub GetOrderSheetDetail_BarCode()
		dim i,sqlStr, sqlsearch

		if frectidxarr="" then exit Sub

		if frectidxarr<>"" then
			sqlsearch = sqlsearch & " and d.masteridx in ("& frectidxarr &")"
		end if
		if FRectMakerid<>"" then
			sqlsearch = sqlsearch & " and d.makerid='" & FRectMakerid & "'"
		end if

		sqlStr = " select"
		sqlStr = sqlStr & " d.idx, d.masteridx, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname"
		sqlStr = sqlStr & " , d.sellcash, d.suplycash, d.buycash, d.baljuitemno, d.realbaljuitemno, d.realitemno"
		sqlStr = sqlStr & " , d.checkitemno, d.regdate, d.updt, d.deldt, d.baljudiv, d.comment, d.makerid"
		sqlStr = sqlStr & " , d.foreign_sellcash, d.foreign_suplycash, c.socname"
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on d.makerid = c.userid"
		sqlStr = sqlStr + " where d.deldt is Null " & sqlsearch
		sqlStr = sqlStr + " order by d.makerid asc, d.itemgubun asc, d.itemid asc, d.itemoption asc"

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Fsuplycash      = rsget("suplycash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fupdt         = rsget("updt")
				FItemList(i).Fdeldt        = rsget("deldt")
				FItemList(i).Fbaljudiv     = rsget("baljudiv")
				FItemList(i).Fcomment      = db2html(rsget("comment"))
				FItemList(i).FMakerid	= rsget("makerid")
				FItemList(i).fforeign_sellcash            = rsget("foreign_sellcash")
				FItemList(i).fforeign_suplycash      = rsget("foreign_suplycash")
				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				FItemList(i).fsocname	= rsget("socname")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public Sub GetLimitCheckSheetDetail()
		dim i,sqlStr
		dim lasteBasedate

		sqlStr = " select d.idx, d.itemgubun, d.itemid, d.itemoption, d.itemname,d.itemoptionname, " + VbCrlf
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.baljuitemno, d.realitemno, d.comment, " + VbCrlf
		sqlStr = sqlStr + " d.makerid, i.sellyn, i.dispyn, i.limityn, i.limitno, i.limitsold, " + VbCrlf
		sqlStr = sqlStr + " IsNULL(s.realstock,0) as realstock, IsNULL(s.ipkumdiv5,0) as ipkumdiv5, IsNULL(s.offconfirmno,0) as offconfirmno, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.offjupno,0) as offjupno, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.ipkumdiv2,0) as ipkumdiv2, IsNULL(s.ipkumdiv4,0) as ipkumdiv4, (case when IsNULL(s.itemid,-1) = -1 then 'Y' else 'N' end) as isnewitem "+ VbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i " + VbCrlf
		sqlStr = sqlStr + " on d.itemgubun='10'" + VbCrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + VbCrlf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s" + VbCrlf
		sqlStr = sqlStr + " on d.itemgubun='10'" + VbCrlf
		sqlStr = sqlStr + " and d.itemgubun=s.itemgubun" + VbCrlf
		sqlStr = sqlStr + " and d.itemid=s.itemid" + VbCrlf
		sqlStr = sqlStr + " and d.itemoption=s.itemoption" + VbCrlf
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and deldt is Null" + VbCrlf
		sqlStr = sqlStr + " order by d.itemgubun, d.itemid, d.itemoption" + VbCrlf
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLimitCheckDetailItem
				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash       = rsget("sellcash")
				FItemList(i).Fsuplycash      = rsget("suplycash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fbaljuitemno   = rsget("baljuitemno")
				FItemList(i).Frealitemno    = rsget("realitemno")
				FItemList(i).Fcomment      = db2html(rsget("comment"))
				FItemList(i).FMakerid	= rsget("makerid")
				FItemList(i).Fsellyn	= rsget("sellyn")
				FItemList(i).Fdispyn	= rsget("dispyn")
				FItemList(i).Flimityn	= rsget("limityn")
				FItemList(i).Flimitno	= rsget("limitno")
				FItemList(i).Flimitsold	= rsget("limitsold")
				''재고파악재고
				FItemList(i).Fcurrno	= rsget("realstock") + rsget("ipkumdiv5") + rsget("offconfirmno")
				''예상재고
				FItemList(i).FMaystockno = FItemList(i).Fcurrno + rsget("ipkumdiv4") + rsget("ipkumdiv2") ''' + rsget("offjupno") (오프라인주문접수건 뺌)
				FItemList(i).FIsNewItem     = rsget("isnewitem")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//admin/fran/jumuninputedit.asp
	public Sub GetOneOrderSheetMaster()
		dim sqlStr

		sqlStr = " select top 1"
		sqlStr = sqlStr & " m.idx,m.targetid,m.baljuid,m.reguser,m.finishuser,m.targetname,m.baljuname,m.regname"
		sqlStr = sqlStr & " ,m.finishname,m.divcode,m.totalsellcash,m.totalsuplycash,isnull(m.totalbuycash,0) as totalbuycash,m.jumunsellcash"
		sqlStr = sqlStr & " ,m.jumunsuplycash,isnull(m.jumunbuycash,0) as jumunbuycash,m.vatinclude,m.regdate,m.updt,m.deldt,m.scheduledate"
		sqlStr = sqlStr & " ,m.beasongdate,m.ipgodate,m.songjangdiv,m.songjangname,m.songjangno,m.baljucode"
		sqlStr = sqlStr & " ,m.statecd,m.comment,m.brandlist,m.alinkcode,m.blinkcode,m.replycomment,m.scheduleipgodate"
		sqlStr = sqlStr & " ,m.sendsms,m.ipkumdate,m.segumdate,m.clinkcode,m.obaljucode,m.workidx,m.shopconfirmuserid"
		sqlStr = sqlStr & " ,m.shopconfirmdate,m.shopconfirmipgodate,m.cwFlag,m.checkusersn,m.rackipgousersn,m.sitename"
		sqlStr = sqlStr & " ,m.currencyUnit,m.foreign_statecd,m.jumunforeign_sellcash,m.jumunforeign_suplycash"
		sqlStr = sqlStr & " ,m.totalforeign_sellcash,m.totalforeign_suplycash,m.accountdiv,m.referip,m.pggubun"
		sqlStr = sqlStr & " ,m.wholesale_ipkumdate,m.paygatetid,m.resultmsg,m.authcode,m.smssenddate"
		sqlStr = sqlStr & " , IsNULL(s.totalsellcash,0) as ipchulsellcash, IsNULL(s.totalsuplycash,0) as ipchulsuplycash, "
		sqlStr = sqlStr & " IsNULL(s.totalbuycash,0) as ipchulbuycash, s.deldt as ipchuldeldt"
		sqlStr = sqlStr & " , p1.username as checkusername, p2.username as rackipgousername "
		sqlStr = sqlStr & " ,ep.reportidx, ep.reportstate, replace(isnull(u.manhp,''),'-','') as manager_hp, u.manemail as manager_email"
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr & " left join [db_storage].[dbo].tbl_acount_storage_master s with (nolock)"
		sqlStr = sqlStr & " 	on m.alinkcode=s.code"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten p1 with (nolock)"
		sqlStr = sqlStr & " 	on m.checkusersn=p1.empno "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten p2 with (nolock)"
		sqlStr = sqlStr & " 	on m.rackipgousersn=p2.empno "
		sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing =1   and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69) "
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u with (nolock)"
		sqlStr = sqlStr & " 	on m.baljuid=u.userid"
		'sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p with (nolock)"
		'sqlStr = sqlStr & " 	on m.baljuid=p.id"
		'sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_group g with (nolock)"
		'sqlStr = sqlStr & " 	on p.groupid=g.groupid"
		sqlStr = sqlStr & " where m.idx=" + CStr(FRectIdx) + ""

		'response.write sqlStr & "<br>"
		'response.end
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then

			set FOneItem = new COrderSheetMasterItem

			FOneItem.fsmssenddate         = rsget("smssenddate")
			FOneItem.fjumunforeign_sellcash         = rsget("jumunforeign_sellcash")
			FOneItem.fjumunforeign_suplycash         = rsget("jumunforeign_suplycash")
			FOneItem.ftotalforeign_sellcash         = rsget("totalforeign_sellcash")
			FOneItem.ftotalforeign_suplycash         = rsget("totalforeign_suplycash")
			FOneItem.fsitename         = rsget("sitename")
			FOneItem.fcurrencyUnit         = rsget("currencyUnit")
			FOneItem.fforeign_statecd            = rsget("foreign_statecd")
			FOneItem.fcwflag         = rsget("cwflag")
			FOneItem.Fidx            = rsget("idx")
			FOneItem.Ftargetid       = rsget("targetid")
			FOneItem.Fbaljuid        = rsget("baljuid")
			FOneItem.Freguser        = rsget("reguser")
			FOneItem.Ffinishuser     = rsget("finishuser")
			FOneItem.Ftargetname     = db2html(rsget("targetname"))
			FOneItem.Fbaljuname      = db2html(rsget("baljuname"))
			FOneItem.Fregname        = db2html(rsget("regname"))
			FOneItem.Ffinishname     = db2html(rsget("finishname"))
			FOneItem.Fdivcode        = rsget("divcode")
			FOneItem.Ftotalsellcash  = rsget("totalsellcash")
			FOneItem.Ftotalsuplycash = rsget("totalsuplycash")
			FOneItem.Ftotalbuycash   = rsget("totalbuycash")
			FOneItem.Fjumunsellcash  = rsget("jumunsellcash")
			FOneItem.Fjumunsuplycash = rsget("jumunsuplycash")
			FOneItem.Fjumunbuycash   = rsget("jumunbuycash")
			FOneItem.Fvatinclude     = rsget("vatinclude")
			FOneItem.Fregdate        = rsget("regdate")
			FOneItem.Fupdt           = rsget("updt")
			FOneItem.Fdeldt          = rsget("deldt")
			FOneItem.Fscheduledate   = rsget("scheduledate")
			FOneItem.Fbeasongdate    = rsget("beasongdate")
			FOneItem.Fipgodate       = rsget("ipgodate")
			FOneItem.Fsongjangdiv    = rsget("songjangdiv")
			FOneItem.Fsongjangname   = db2html(rsget("songjangname"))
			FOneItem.Fsongjangno     = rsget("songjangno")
			FOneItem.Fbaljucode      = rsget("baljucode")
			FOneItem.Fstatecd        = rsget("statecd")
			FOneItem.Fcomment        = db2html(rsget("comment"))
			FOneItem.FBrandList	    = rsget("brandlist")
			FOneItem.Fscheduleipgodate = rsget("scheduleipgodate")
			FOneItem.Freplycomment = db2html(rsget("replycomment"))
			FOneItem.Fsendsms		= rsget("sendsms")
			FOneItem.Falinkcode		= rsget("alinkcode")
			FOneItem.Fblinkcode		= rsget("blinkcode")
			FOneItem.Fipchulsellcash		= rsget("ipchulsellcash")
			FOneItem.Fipchulsuplycash		= rsget("ipchulsuplycash")
			FOneItem.Fipchulbuycash		= rsget("ipchulbuycash")
			FOneItem.Fipchuldeldt		= rsget("ipchuldeldt")
			FOneItem.fshopconfirmuserid		= rsget("shopconfirmuserid")
			FOneItem.fshopconfirmdate		= rsget("shopconfirmdate")
			FOneItem.fshopconfirmipgodate		= rsget("shopconfirmipgodate")

			FOneItem.Fcheckusersn		= rsget("checkusersn")
			FOneItem.Frackipgousersn	= rsget("rackipgousersn")

			FOneItem.Fcheckusername		= rsget("checkusername")
			FOneItem.Frackipgousername	= rsget("rackipgousername")

			FOneItem.Freportidx      = rsget("reportidx")
			FOneItem.Freportstate    = rsget("reportstate")
			FOneItem.fmanager_hp     = db2html(rsget("manager_hp"))
			FOneItem.fmanager_email     = db2html(rsget("manager_email"))
		end if
		rsget.Close

        if (FResultCount > 0) then
            if IsNull(FOneItem.Freportidx) then
		        sqlStr = " select top 1 pm.idx as ppMasterIdx, ep.reportidx as reportIdx, ep.reportstate "
		        sqlStr = sqlStr + " from "
		        sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_pp_product_link] l "
		        sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_pp_product_master] pm on l.ppMasterIdx = pm.idx "
                sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep on l.ppMasterIdx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104) "
		        sqlStr = sqlStr + " where "
		        sqlStr = sqlStr + " 	1 = 1 "
		        sqlStr = sqlStr + " 	and l.linkidx = " & FRectIdx
		        sqlStr = sqlStr + " 	and l.linkType = 'JUMUN' "
		        sqlStr = sqlStr + " 	and l.deldt is NULL "
		        sqlStr = sqlStr + " 	and pm.deldt is NULL "
                ''response.write sqlStr
		        rsget.Open sqlStr, dbget, 1
		        if Not rsget.Eof then
                    FOneItem.FppReportidx    = rsget("reportidx")
                    FOneItem.FppMasterIdx    = rsget("ppMasterIdx")
                    FOneItem.FppReportstate  = rsget("reportstate")
                end if
                rsget.Close
            end if
        end if
	end Sub

	'//admin/fran/jumunlist.asp
	public Sub GetOrderSheetListByBrand()
		dim i,sqlStr, sqlsearch

		if FRectBaljuCode<>"" then
			sqlsearch = sqlsearch + " and m.baljucode='" + FRectBaljuCode + "'"
		end if

		if FRectDivCodeArr<>"" then
			sqlsearch = sqlsearch + " and m.divcode  in " + FRectDivCodeArr + ""
		end if

		if FRectDivCode<>"" then
			sqlsearch = sqlsearch + " and m.divcode='" + FRectDivCode + "'"
		end if

		if FRectDivCodeUnder<>"" then
			sqlsearch = sqlsearch + " and m.divcode<'" + FRectDivCodeUnder + "'"
		end if

		if FRectStatecd<>"" then
			if FRectStatecd="foreign0" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=0"
			elseif FRectStatecd="foreign3" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=3"
			elseif FRectStatecd="foreign7" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=7"
			elseif FRectStatecd=" " then
				sqlsearch = sqlsearch + " and m.statecd='" + FRectStatecd + "' and m.foreign_statecd is null"
			elseif FRectStatecd="before1" then
				sqlsearch = sqlsearch + " and m.statecd < '1' "
			else
				sqlsearch = sqlsearch + " and m.statecd='" + FRectStatecd + "'"
			end if
		end if

		if FRectBaljuId<>"" then
			sqlsearch = sqlsearch + " and m.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectTargetid<>"" then
			sqlsearch = sqlsearch + " and m.targetid='" + FRectTargetid + "'"
		end if

		if FRectStartDate<>"" then
			sqlsearch = sqlsearch + " and m.scheduledate>='" + FRectStartDate + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and m.scheduledate<'" + FRectEndDate + "'"
		end if

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid ='" + FRectMakerid + "'"
		end if

		if FRectNotIpgoOnly<>"" then
			sqlsearch = sqlsearch + " and m.alinkcode is null"
		end if

		if FRectIdxarr<>"" then
			sqlsearch = sqlsearch + " and m.idx in(" + CStr(idxarr) + ")"
		end if

		if FRectMinusOnly<>"" then
			sqlsearch = sqlsearch + " and m.totalsellcash<0"
		end if

		if FRectShopDiv="f87" then
            sqlsearch = sqlsearch + " and Left(baljuid,12)='streetshop87'"
		elseif FRectShopDiv="f" then
			sqlsearch = sqlsearch + " and Left(baljuid,11)='streetshop8'"
		elseif FRectShopDiv="j" then
			sqlsearch = sqlsearch + " and Left(baljuid,11)<>'streetshop8'"
		elseif Len(FRectDivGubun)=3 then
		    sqlsearch = sqlsearch + " and divcode ='" + FRectDivGubun + "'"
		end if

        if (FRectReOrderOnly<>"") then
            sqlsearch = sqlsearch + " and Left(baljucode,2)='RJ'"
        end if

		if FRectDivGubun="j" then
			sqlsearch = sqlsearch + " and m.divcode in ('101','111')"
		elseif FRectDivGubun="p" then
			sqlsearch = sqlsearch + " and m.divcode in ('121','201')"
		elseif FRectDivGubun="c" then
			sqlsearch = sqlsearch + " and m.divcode ='131'"
		elseif Len(FRectDivGubun)=3 then
		    sqlsearch = sqlsearch + " and m.divcode ='" + FRectDivGubun + "'"
		end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlsearch = sqlsearch + " 	and m.baljuid not in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
			else
				sqlsearch = sqlsearch + " 	and m.baljuid in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlsearch = sqlsearch + " and d.itemgubun ='" & FRectItemGubun & "'"
		end if

		if (FRectItemID <> "") then
			sqlsearch = sqlsearch + " and d.itemid = " & FRectItemID
		end if

		if (FRectItemOption <> "") then
			sqlsearch = sqlsearch + " and d.itemoption ='" & FRectItemOption & "'"
		end if

		sqlStr = " select count(distinct m.idx) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx " & sqlsearch
		sqlStr = sqlStr + " and m.deldt is Null"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select distinct top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.idx,m.baljuid,m.targetid,m.targetname,m.baljuname,m.divcode,"
		sqlStr = sqlStr + " m.totalsellcash,m.totalsuplycash,m.totalbuycash,m.jumunsellcash,m.jumunsuplycash,m.jumunbuycash,"
		sqlStr = sqlStr + " m.scheduledate,m.baljucode,m.statecd,m.beasongdate,m.songjangname,m.songjangno,m.songjangdiv,"
		sqlStr = sqlStr + " m.regdate,m.brandlist,m.scheduleipgodate,m.sendsms,m.regname,m.reguser,m.finishname,m.ipgodate,"
		sqlStr = sqlStr + " m.ipkumdate,m.segumdate,m.alinkcode,m.blinkcode, m.currencyUnit, m.foreign_statecd, m.sitename"
		sqlStr = sqlStr + " , m.smssenddate"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx " & sqlsearch
		sqlStr = sqlStr + " and m.deldt is Null"
		sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem
				FItemList(i).fsmssenddate          = rsget("smssenddate")
				FItemList(i).fsitename          = rsget("sitename")
				FItemList(i).FcurrencyUnit          = rsget("currencyUnit")
				FItemList(i).Fforeign_statecd        = rsget("foreign_statecd")
				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fbaljuid        = rsget("baljuid")
				FItemList(i).Ftargetid       = rsget("targetid")
				FItemList(i).Ftargetname     = db2html(rsget("targetname"))
				FItemList(i).Fbaljuname     = db2html(rsget("baljuname"))
				FItemList(i).Fdivcode       = rsget("divcode")
				FItemList(i).Ftotalsellcash    = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash    = rsget("totalsuplycash")
				FItemList(i).Ftotalbuycash    = rsget("totalbuycash")
				FItemList(i).Fjumunsellcash = rsget("jumunsellcash")
				FItemList(i).Fjumunsuplycash	= rsget("jumunsuplycash")
				FItemList(i).Fjumunbuycash	= rsget("jumunbuycash")
				FItemList(i).Fscheduledate		= rsget("scheduledate")
				FItemList(i).Fbaljucode		= rsget("baljucode")
				FItemList(i).Fstatecd		= rsget("statecd")
				FItemList(i).Fbeasongdate		= rsget("beasongdate")
				FItemList(i).Fsongjangname		= db2html(rsget("songjangname"))
				FItemList(i).Fsongjangno	= rsget("songjangno")
				FItemList(i).Fsongjangdiv	= rsget("songjangdiv")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FBrandList	    = rsget("brandlist")
				FItemList(i).Fscheduleipgodate = rsget("scheduleipgodate")
				FItemList(i).Fsendsms		= rsget("sendsms")
				FItemList(i).Fregname		= db2html(rsget("regname"))
				FItemList(i).Freguser		= rsget("reguser")
				FItemList(i).Ffinishname	= db2html(rsget("finishname"))
				FItemList(i).Fipgodate		= rsget("ipgodate")
				FItemList(i).Fipkumdate		= rsget("ipkumdate")
				FItemList(i).Fsegumdate		= rsget("segumdate")
				FItemList(i).Falinkcode		= rsget("alinkcode")
				FItemList(i).Fblinkcode		= rsget("blinkcode")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

		if frecttotalyn="Y" then
			sqlStr = " select"
			sqlStr = sqlStr & " sum(m.jumunsellcash) as jumunsellcash, sum(m.jumunsuplycash) as jumunsuplycash, sum(m.totalsuplycash) as totalsuplycash"
			sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d"
			sqlStr = sqlStr & " 	on m.idx=d.masteridx"
			sqlStr = sqlStr & " where m.deldt is Null " & sqlsearch

			'response.write sqlStr &"<br>"
			rsget.pagesize = FPageSize
			rsget.Open sqlStr, dbget, 1

			if not rsget.EOF  then
				total_jumunsellcash = rsget("jumunsellcash")
				total_jumunsuplycash = rsget("jumunsuplycash")
				total_totalsuplycash = rsget("totalsuplycash")
			else
				total_jumunsellcash = 0
				total_jumunsuplycash = 0
				total_totalsuplycash = 0
			end if
			rsget.close
		end if
	end Sub

	'//admin/fran/jumunlist.asp
	public Sub GetOrderSheetList()
		dim i,sqlStr , sqlsearch, tmpStr

		if frectmakerid <> "" then
			sqlsearch = sqlsearch + " and m.brandlist like '%" + CStr(frectmakerid) + "%'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.regdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.regdate<'" + CStr(FRectToDate) + "'"
			end if
		'//입고일 기준
		elseif frectdatefg = "ipgo" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.ipgodate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.ipgodate<'" + CStr(FRectToDate) + "'"
			end if
		end if

		if FRectBaljuCode<>"" then
			if (InStr(FRectBaljuCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBaljuCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and m.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and m.baljucode='" + FRectBaljuCode + "'"
			end if
		end if

		if FRectDivCodeArr<>"" then
			sqlsearch = sqlsearch + " and m.divcode  in " + FRectDivCodeArr + ""
		end if

		if FRectDivCode<>"" then
			sqlsearch = sqlsearch + " and m.divcode='" + FRectDivCode + "'"
		end if

		if FRectDivCodeUnder<>"" then
			sqlsearch = sqlsearch + " and m.divcode<'" + FRectDivCodeUnder + "'"
		end if

		if FRectStatecd<>"" then
			if FRectStatecd="foreign0" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=0"
			elseif FRectStatecd="foreign3" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=3"
			elseif FRectStatecd="foreign7" then
				sqlsearch = sqlsearch + " and m.statecd=' ' and m.foreign_statecd=7"
			elseif FRectStatecd=" " then
				sqlsearch = sqlsearch + " and m.statecd='" + FRectStatecd + "' and m.foreign_statecd is null"
			elseif FRectStatecd="preorder" then		'//기주문상태
				sqlsearch = sqlsearch + " and m.statecd<9 and m.scheduledate >= dateadd(m,-2,getdate())"
			elseif FRectStatecd="before1" then
				sqlsearch = sqlsearch + " and m.statecd < '1' "
			else
				sqlsearch = sqlsearch + " and m.statecd='" + FRectStatecd + "'"
			end if
		end if

		if FRectReportState <> "" then
			if FRectReportState ="7" then
				sqlsearch = sqlsearch + " and ep.reportstate in ('7','8','9')"
			else
					sqlsearch = sqlsearch + " and ep.reportstate='" + FRectReportState + "'"
			end if
		end if

		if FRectbaljunum<>"" then
			sqlsearch = sqlsearch + " and b.baljunum=" + FRectbaljunum + " "
		end if

		if FRectBaljuId<>"" then
			sqlsearch = sqlsearch + " and m.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectTargetid<>"" then
			sqlsearch = sqlsearch + " and m.targetid='" + FRectTargetid + "'"
		end if

		if FRectStartDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate>='" + FRectStartDate + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate<'" + FRectEndDate + "'"
		end if

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and m.brandlist + ',' like '%" + FRectMakerid + ",%'"
		end if

		if FRectNotIpgoOnly<>"" then
			sqlsearch = sqlsearch + " and m.alinkcode is null"
		end if

		if FRectBLinkCode<>"" then
			if (InStr(FRectBLinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBLinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and m.blinkcode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and m.blinkcode='" + FRectBLinkCode + "'"
			end if
		end if

		if FRectIdxarr<>"" then
			sqlsearch = sqlsearch + " and m.idx in(" + CStr(idxarr) + ")"
		end if

        if FRectMinusOnly = "Y" then
            sqlsearch = sqlsearch + " and m.totalsellcash < 0"
        elseif FRectMinusOnly = "N" then
            sqlsearch = sqlsearch + " and m.totalsellcash >= 0"
		elseif FRectMinusOnly<>"" then
			sqlsearch = sqlsearch + " and m.totalsellcash < 0"
		end if

        if FRectShopDiv="f87" then
            sqlsearch = sqlsearch + " and Left(m.baljuid,12)='streetshop87'"
		elseif FRectShopDiv="f" then
			sqlsearch = sqlsearch + " and Left(m.baljuid,11)='streetshop8'"
		elseif FRectShopDiv="j" then
			sqlsearch = sqlsearch + " and Left(m.baljuid,11)<>'streetshop8'"
		elseif Len(FRectDivGubun)=3 then
		    sqlsearch = sqlsearch + " and m.divcode ='" + FRectDivGubun + "'"
		end if

		if FRectDivGubun="j" then
			sqlsearch = sqlsearch + " and m.divcode in ('101','111')"
		elseif FRectDivGubun="p" then
			sqlsearch = sqlsearch + " and m.divcode in ('121','201')"
		elseif FRectDivGubun="c" then
			sqlsearch = sqlsearch + " and m.divcode ='131'"
		elseif Len(FRectDivGubun)=3 then
		    sqlsearch = sqlsearch + " and m.divcode ='" + FRectDivGubun + "'"
		end if

        if (FRectReOrderOnly<>"") then
            sqlsearch = sqlsearch + " and Left(m.baljucode,2)='RJ'"
        end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlsearch = sqlsearch + " 	and m.baljuid not in (select id from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
			else
				sqlsearch = sqlsearch + " 	and m.baljuid in (select id from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if
		if FRectitemgubun<>"" and FRectitemid<>"" and FRectitemoption<>"" then
			sqlsearch = sqlsearch + " 	and m.idx in ("
			sqlsearch = sqlsearch + " 		select masteridx"
			sqlsearch = sqlsearch + " 		from [db_storage].[dbo].tbl_ordersheet_detail with (nolock)"
			sqlsearch = sqlsearch + " 		where itemgubun='"& FRectitemgubun &"'"
			sqlsearch = sqlsearch + " 		and itemid="& FRectitemid &""
			sqlsearch = sqlsearch + " 		and itemoption='"& FRectitemoption &"'"
			sqlsearch = sqlsearch + " 		and isnull(deldt,'')=''"
			sqlsearch = sqlsearch + " 	)"
		elseif FRectitemid<>"" then
			sqlsearch = sqlsearch + " 	and m.idx in ("
			sqlsearch = sqlsearch + " 		select masteridx"
			sqlsearch = sqlsearch + " 		from [db_storage].[dbo].tbl_ordersheet_detail with (nolock)"
			sqlsearch = sqlsearch + " 		where itemid="& FRectitemid &" "
			sqlsearch = sqlsearch + " 		and isnull(deldt,'')=''"
			sqlsearch = sqlsearch + " 	)"
		end if
		if (FRectBrandPurchaseType<>"") then
		    sqlsearch = sqlsearch + " and p.purchaseType="&FRectBrandPurchaseType&""
		end if

		if (trim(FRectSearchField) <> "" and trim(FRectSearchText)<>"") then
			if trim(FRectSearchField)="socname" then
				sqlsearch = sqlsearch + " and m.targetid in ( "
				sqlsearch = sqlsearch + " 	select distinct p1.id "
				sqlsearch = sqlsearch + " 	from "
				sqlsearch = sqlsearch + " 		db_partner.dbo.tbl_partner p1 with (nolock)"
				sqlsearch = sqlsearch + " 		join [db_partner].[dbo].tbl_partner_group g1 with (nolock)"
				sqlsearch = sqlsearch + " 		on "
				sqlsearch = sqlsearch + " 			p1.groupid = g1.groupid "
				sqlsearch = sqlsearch + " 	where "
				sqlsearch = sqlsearch + " 		1 = 1 "
				sqlsearch = sqlsearch + " 		and g1.company_name like '%" & trim(FRectSearchText) & "%' "
				sqlsearch = sqlsearch + " 		and p1.isusing = 'Y' "
				sqlsearch = sqlsearch + " ) "
			elseif trim(FRectSearchField)="socno" then
				sqlsearch = sqlsearch + " and m.targetid in ( "
				sqlsearch = sqlsearch + " 	select distinct p1.id "
				sqlsearch = sqlsearch + " 	from "
				sqlsearch = sqlsearch + " 		db_partner.dbo.tbl_partner p1 with (nolock)"
				sqlsearch = sqlsearch + " 		join [db_partner].[dbo].tbl_partner_group g1 with (nolock)"
				sqlsearch = sqlsearch + " 		on "
				sqlsearch = sqlsearch + " 			p1.groupid = g1.groupid "
				sqlsearch = sqlsearch + " 	where "
				sqlsearch = sqlsearch + " 		1 = 1 "
				sqlsearch = sqlsearch + " 		and replace(g1.company_no,'-','') = '" & trim(replace(FRectSearchText,"-","")) & "' "
				sqlsearch = sqlsearch + " 		and p1.isusing = 'Y' "
				sqlsearch = sqlsearch + " ) "
			else
				sqlsearch = sqlsearch + " and " & trim(FRectSearchField) & "='" & trim(FRectSearchText) & "'"
			end if
		end if

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
	    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.targetid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on m.baljuid=pp.id"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 	on b.baljucode = m.baljucode "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p1 with (nolock)"
		sqlStr = sqlStr + " 	on m.checkusersn=p1.empno "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p2 with (nolock)"
		sqlStr = sqlStr + " 	on m.rackipgousersn=p2.empno "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing =1   and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69) "
		sqlStr = sqlStr + " where m.deldt is Null " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top 1 (case when count(l.workSecond)<1 then '00:00:00' else left(convert(varchar, DATEADD(second, IsNull(sum(IsNull(l.workSecond,0))/ count(l.workSecond),0), 0), 114),8) end) as AverageWorkSecond"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
	    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.targetid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on m.baljuid=pp.id"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 	on b.baljucode = m.baljucode "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p1 with (nolock)"
		sqlStr = sqlStr + " 	on m.checkusersn=p1.empno "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p2 with (nolock)"
		sqlStr = sqlStr + " 	on m.rackipgousersn=p2.empno "
		sqlStr = sqlStr + " left join [db_log].[dbo].[tbl_logics_act_log] l with (nolock)"
		sqlStr = sqlStr + " 	on m.checkusersn = l.empno and l.actgubun = 'onLineIpgoCheck' and l.refcode = m.baljucode "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing =1   and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69) "
		sqlStr = sqlStr + " where m.deldt is Null " & sqlsearch

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF  then
			FAverageWorkSecond	= rsget("AverageWorkSecond")
		end if
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx,m.targetid,m.baljuid,m.reguser,m.finishuser,m.targetname,m.baljuname,m.regname"
		sqlStr = sqlStr & " ,m.finishname,m.divcode,m.totalsellcash,m.totalsuplycash,isnull(m.totalbuycash,0) as totalbuycash,m.jumunsellcash"
		sqlStr = sqlStr & " ,m.jumunsuplycash,isnull(m.jumunbuycash,0) as jumunbuycash,m.vatinclude,m.regdate,m.updt,m.deldt,m.scheduledate"
		sqlStr = sqlStr & " ,m.beasongdate,m.ipgodate,m.songjangdiv,m.songjangname,m.songjangno,m.baljucode"
		sqlStr = sqlStr & " ,m.statecd,m.comment,m.brandlist,m.alinkcode,m.blinkcode,m.replycomment,m.scheduleipgodate"
		sqlStr = sqlStr & " ,m.sendsms,m.ipkumdate,m.segumdate,m.clinkcode,m.obaljucode,m.workidx,m.shopconfirmuserid"
		sqlStr = sqlStr & " ,m.shopconfirmdate,m.shopconfirmipgodate,m.cwFlag,m.checkusersn,m.rackipgousersn,m.sitename"
		sqlStr = sqlStr & " ,m.currencyUnit,m.foreign_statecd,m.jumunforeign_sellcash,m.jumunforeign_suplycash"
		sqlStr = sqlStr & " ,m.totalforeign_sellcash,m.totalforeign_suplycash,m.accountdiv,m.referip,m.pggubun"
		sqlStr = sqlStr & " ,m.wholesale_ipkumdate,m.paygatetid,m.resultmsg,m.authcode,m.smssenddate"
		sqlStr = sqlStr + " , b.baljudate, b.baljunum "
		sqlStr = sqlStr + " , p1.username as checkusername, p2.username as rackipgousername, p.purchasetype, pp.tplcompanyid, (case when l.workSecond is NULL then '' else left(convert(varchar, DATEADD(second, IsNull(workSecond,0), 0), 114),8) end) as workSecond"
		sqlStr = sqlStr + " , ep.reportidx, ep.reportstate"
		sqlStr = sqlStr + " , ("
		sqlStr = sqlStr + " 	select top 1"
		sqlStr = sqlStr + " 	pl.ppmasteridx"
		sqlStr = sqlStr + " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		sqlStr = sqlStr + " 		on pm.idx=pl.ppMasterIdx"
		sqlStr = sqlStr + " 	where pm.deldt is null"
		sqlStr = sqlStr + " 	and pl.deldt is null"
		sqlStr = sqlStr + " 	and m.idx=pl.linkIdx"
		sqlStr = sqlStr + " 	order by pl.ppmasteridx desc"
		sqlStr = sqlStr + " 	) as ppmasteridx"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
	    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.targetid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on m.baljuid=pp.id"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 	on b.baljucode = m.baljucode "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p1 with (nolock)"
		sqlStr = sqlStr + " 	on m.checkusersn=p1.empno "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten p2 with (nolock)"
		sqlStr = sqlStr + " 	on m.rackipgousersn=p2.empno "
		sqlStr = sqlStr + " left join [db_log].[dbo].[tbl_logics_act_log] l with (nolock)"
		sqlStr = sqlStr + " 	on m.checkusersn = l.empno and l.actgubun = 'onLineIpgoCheck' and l.refcode = m.baljucode "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing =1   and (ep.edmsidx = 65 or ep.edmsidx = 68 or ep.edmsidx = 69) "
		sqlStr = sqlStr + " where m.deldt is Null " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

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

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem
				FItemList(i).fsmssenddate          = rsget("smssenddate")
				FItemList(i).Ftplcompanyid          = rsget("tplcompanyid")
				FItemList(i).fpurchasetype          = rsget("purchasetype")
				FItemList(i).fsitename          = rsget("sitename")
				FItemList(i).fcurrencyUnit          = rsget("currencyUnit")
				FItemList(i).fforeign_statecd        = rsget("foreign_statecd")
				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fbaljuid        = rsget("baljuid")
				FItemList(i).Ftargetid       = rsget("targetid")
				FItemList(i).Ftargetname     = db2html(rsget("targetname"))
				FItemList(i).Fbaljuname     = db2html(rsget("baljuname"))
				FItemList(i).Fdivcode       = rsget("divcode")
				FItemList(i).Ftotalsellcash    = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash    = rsget("totalsuplycash")
				FItemList(i).Ftotalbuycash    = rsget("totalbuycash")
				FItemList(i).Fjumunsellcash = rsget("jumunsellcash")
				FItemList(i).Fjumunsuplycash	= rsget("jumunsuplycash")
				FItemList(i).Fjumunbuycash	= rsget("jumunbuycash")
				FItemList(i).Fscheduledate		= rsget("scheduledate")
				FItemList(i).Fbaljucode		= rsget("baljucode")
				FItemList(i).Fstatecd		= rsget("statecd")
				FItemList(i).Fbeasongdate		= rsget("beasongdate")
				FItemList(i).Fsongjangname		= db2html(rsget("songjangname"))
				FItemList(i).Fsongjangno	= rsget("songjangno")
				FItemList(i).Fsongjangdiv	= rsget("songjangdiv")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FBrandList	    = rsget("brandlist")
				FItemList(i).Fscheduleipgodate = rsget("scheduleipgodate")
				FItemList(i).Fsendsms		= rsget("sendsms")
				FItemList(i).Fregname		= db2html(rsget("regname"))
				FItemList(i).Freguser		= rsget("reguser")
				FItemList(i).Ffinishname	= db2html(rsget("finishname"))
				FItemList(i).Fcomment		= db2html(rsget("comment"))
				FItemList(i).Fipgodate		= rsget("ipgodate")
				FItemList(i).Fipkumdate		= rsget("ipkumdate")
				FItemList(i).Fsegumdate		= rsget("segumdate")
				FItemList(i).Falinkcode		= rsget("alinkcode")
				FItemList(i).Fblinkcode		= rsget("blinkcode")

				FItemList(i).Fbaljudate		= rsget("baljudate")
				FItemList(i).Fbaljunum		= rsget("baljunum")

				FItemList(i).Fcheckusersn		= rsget("checkusersn")
				FItemList(i).Frackipgousersn	= rsget("rackipgousersn")

				FItemList(i).Fcheckusername		= rsget("checkusername")
				FItemList(i).Frackipgousername	= rsget("rackipgousername")

				FItemList(i).FworkSecond	= rsget("workSecond")

				FItemList(i).Freportidx	= rsget("reportidx")
				FItemList(i).Freportstate	= rsget("reportstate")

				FItemList(i).Fjumunforeign_sellcash	= rsget("jumunforeign_sellcash")
				FItemList(i).Fjumunforeign_suplycash	= rsget("jumunforeign_suplycash")
				FItemList(i).Ftotalforeign_suplycash	= rsget("totalforeign_suplycash")

                FItemList(i).FppMasterIdx	= rsget("ppMasterIdx")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

		if frecttotalyn="Y" then
			sqlStr = " select"
			sqlStr = sqlStr & " sum(m.jumunsellcash) as jumunsellcash, sum(m.jumunsuplycash) as jumunsuplycash, sum(m.totalsuplycash) as totalsuplycash"
			sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master m "

			if (FRectBrandPurchaseType<>"") then
			    sqlStr = sqlStr & " Join db_partner.dbo.tbl_partner p"
			    sqlStr = sqlStr & " 	on m.targetid=p.id"
			    sqlStr = sqlStr & " 	and p.purchaseType="&FRectBrandPurchaseType&""
			end if

			sqlStr = sqlStr & " left join [db_storage].[dbo].tbl_shopbalju b "
			sqlStr = sqlStr & " 	on b.baljucode = m.baljucode "
			sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten p1"
			sqlStr = sqlStr & " 	on m.checkusersn=p1.empno "
			sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten p2"
			sqlStr = sqlStr & " 	on m.rackipgousersn=p2.empno "
			sqlStr = sqlStr & " where m.deldt is Null " & sqlsearch

			'response.write sqlStr &"<br>"
			rsget.Open sqlStr, dbget, 1

			if not rsget.EOF  then
				total_jumunsellcash = rsget("jumunsellcash")
				total_jumunsuplycash = rsget("jumunsuplycash")
				total_totalsuplycash = rsget("totalsuplycash")
			else
				total_jumunsellcash = 0
				total_jumunsuplycash = 0
				total_totalsuplycash = 0
			end if
			rsget.close
		end if
	end Sub

	public Sub GetOrderSheetListWithBrand()
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master"
		sqlStr = sqlStr + " where deldt is Null"
		if FRectDivCode<>"" then
			sqlStr = sqlStr + " and divcode='" + FRectDivCode + "'"
		end if

		if FRectStatecd<>"" then
			sqlStr = sqlStr + " and statecd='" + FRectStatecd + "'"
		end if

		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and baljuid='" + FRectBaljuId + "'"
		end if

		if FRectTargetid<>"" then
			sqlStr = sqlStr + " and targetid='" + FRectTargetid + "'"
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and scheduledate>='" + FRectStartDate + "'"
		end if

		if FRectEndDate<>"" then
			sqlStr = sqlStr + " and scheduledate<'" + FRectEndDate + "'"
		end if


		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master"
		sqlStr = sqlStr + " where deldt is Null"
		if FRectDivCode<>"" then
			sqlStr = sqlStr + " and divcode='" + FRectDivCode + "'"
		end if

		if FRectStatecd<>"" then
			sqlStr = sqlStr + " and statecd='" + FRectStatecd + "'"
		end if

		if FRectBaljuId<>"" then
			sqlStr = sqlStr + " and baljuid='" + FRectBaljuId + "'"
		end if

		if FRectTargetid<>"" then
			sqlStr = sqlStr + " and targetid='" + FRectTargetid + "'"
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and scheduledate>='" + FRectStartDate + "'"
		end if

		if FRectEndDate<>"" then
			sqlStr = sqlStr + " and scheduledate<'" + FRectEndDate + "'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetMasterItem

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fbaljuid        = rsget("baljuid")
				FItemList(i).Ftargetid       = rsget("targetid")
				FItemList(i).Ftargetname     = db2html(rsget("targetname"))
				FItemList(i).Fbaljuname     = db2html(rsget("baljuname"))
				FItemList(i).Fdivcode       = rsget("divcode")
				FItemList(i).Ftotalsellcash    = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash    = rsget("totalsuplycash")
				FItemList(i).Fjumunsellcash = rsget("jumunsellcash")
				FItemList(i).Fjumunsuplycash	= rsget("jumunsuplycash")
				FItemList(i).Fscheduledate		= rsget("scheduledate")
				FItemList(i).Fbaljucode		= rsget("baljucode")
				FItemList(i).Fstatecd		= rsget("statecd")
				FItemList(i).Fbeasongdate		= rsget("beasongdate")
				FItemList(i).Fsongjangname		= db2html(rsget("songjangname"))
				FItemList(i).Fregdate		= rsget("regdate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	' 사용안하는듯
	' public function MakeUpcheJumun()
	' 	dim i,sqlStr
	' 	dim iid, baljucode, targetname

	' 	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"
	' 	rsget.Open sqlStr,dbget,1,3
	' 	rsget.AddNew
	' 	rsget("targetid") = FRectTargetid
	' 	rsget("baljuid") = FRectBaljuid
	' 	rsget("baljuname") = FRectBaljuname
	' 	rsget("reguser") = FRectReguser
	' 	rsget("regname") = FRectRegname
	' 	rsget("divcode") = FRectdivcode
	' 	rsget("vatinclude") = "Y"
	' 	rsget("scheduledate") = FRectScheduledate
	' 	rsget("statecd") = "0"
	' 	rsget("brandlist") = FRectMakerid
	' 	rsget("comment") = FRectComment

	' 	rsget.update
	' 	iid = rsget("idx")
	' 	rsget.close

	' 	baljucode = "UJ" + Format00(6,Right(CStr(iid),6))

	' 	if FRectTargetid="10x10" then
	' 		targetname = "텐바이텐"
	' 	else
	' 		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
	' 		sqlStr = sqlStr + " where userid='" + FRectTargetid + "'"
	' 		rsget.Open sqlStr, dbget, 1
	' 		if Not rsget.Eof then
	' 			targetname = rsget("socname_kor")
	' 		end if
	' 		rsget.close
	' 	end if

	' 	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	' 	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	' 	sqlStr = sqlStr + " ,targetname='" + html2db(targetname) + "'" + VbCrlf
	' 	sqlStr = sqlStr + " where idx=" + CStr(iid)
	' 	rsget.Open sqlStr, dbget, 1

	' 	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	' 	sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
	' 	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	' 	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
	' 	sqlStr = sqlStr + " select " + CStr(iid)  + ", d.itemgubun, d.makerid, d.itemid, d.itemoption," + vbCrlf
	' 	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.sellcash, d.suplycash, d.buycash," + vbCrlf
	' 	sqlStr = sqlStr + " sum(d.baljuitemno),sum(d.baljuitemno),'0'"  + vbCrlf
	' 	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
	' 	sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_ordersheet_detail d"
	' 	''--오프 or 업체배송 상품만?.
	' 	sqlStr = sqlStr + " where m.idx<>0 "
	' 	sqlStr = sqlStr + " and m.idx=d.masteridx"
	' 	sqlStr = sqlStr + " and m.idx in (" + FRectIdxArr + ")"
	' 	sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
	' 	sqlStr = sqlStr + " group by d.itemgubun, d.makerid, d.itemid, d.itemoption," + vbCrlf
	' 	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.sellcash, d.suplycash, d.buycash" + vbCrlf
	' 	rsget.Open sqlStr, dbget, 1

	' 	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	' 	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	' 	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	' 	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	' 	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	' 	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	' 	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	' 	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	' 	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	' 	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	' 	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	' 	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	' 	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	' 	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	' 	sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
	' 	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	' 	sqlStr = sqlStr + " ) as T" + vbCrlf
	' 	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)

	' 	rsget.Open sqlStr, dbget, 1

	' 	MakeUpcheJumun = iid
	' end function

	public function MakeUpcheJumunNew()
		dim i,sqlStr
		dim iid, baljucode, targetname, orgbaljuname
		dim itemgubunarr, itemidarr, itemoptionarr, ordernoarr
		dim cnt
		dim comment

		sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("targetid") = FRectTargetid
		rsget("baljuid") = FRectBaljuid
		rsget("baljuname") = FRectBaljuname
		rsget("reguser") = FRectReguser
		rsget("regname") = FRectRegname
		rsget("divcode") = FRectdivcode
		rsget("vatinclude") = "Y"
		rsget("scheduledate") = FRectScheduledate
		rsget("statecd") = "0"
		rsget("brandlist") = FRectMakerid
		rsget("comment") = ""

		rsget.update
		iid = rsget("idx")
		rsget.close

		baljucode = "UJ" + Format00(6,Right(CStr(iid),6))

		if FRectTargetid="10x10" then
			targetname = "텐바이텐"
		else
			sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
			sqlStr = sqlStr + " where userid='" + FRectTargetid + "'"
			rsget.Open sqlStr, dbget, 1
			if Not rsget.Eof then
				targetname = rsget("socname_kor")
			end if
			rsget.close
		end if

        '' FRectOrgBaljuCode 가 제대로 넘어 오지 않음..
		sqlStr = " select distinct baljucode, baljuname from [db_storage].[dbo].tbl_ordersheet_master "
		sqlStr = sqlStr + " where baljucode in (" + CStr(FRectOrgBaljuCode) + ")"
		'response.write sqlStr + "<br><br>"
		rsget.Open sqlStr, dbget, 1

		comment = ""
		if Not rsget.Eof then
			do until rsget.eof
				if (comment = "") then
					comment = "원주문 : " + rsget("baljucode") + "[" + rsget("baljuname") + "]"
				else
					comment = comment + ", " + rsget("baljucode") + "[" + rsget("baljuname") + "]"
				end if
				rsget.MoveNext
			loop
		end if
		rsget.close

		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
		sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
		sqlStr = sqlStr + " ,targetname='" + html2db(targetname) + "'" + VbCrlf
		sqlStr = sqlStr + " ,comment='" + html2db(comment) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(iid)
		rsget.Open sqlStr, dbget, 1

		'원주문코드 입력
		sqlStr = " update d" + VbCrlf
		sqlStr = sqlStr + " set " + VbCrlf
		sqlStr = sqlStr + " 	d.upcheorderlinkcode = '" + baljucode + "' " + VbCrlf
		sqlStr = sqlStr + " from " + VbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m " + VbCrlf
		sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d " + VbCrlf
		sqlStr = sqlStr + " 	on " + VbCrlf
		sqlStr = sqlStr + " 		m.idx = d.masteridx " + VbCrlf
		sqlStr = sqlStr + " where " + VbCrlf
		sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
		sqlStr = sqlStr + " 	and m.baljucode in (" + CStr(FRectOrgBaljuCode) + ") " + VbCrlf
		sqlStr = sqlStr + " 	and m.statecd in ('0', '1') " + VbCrlf
		sqlStr = sqlStr + " 	and d.makerid = '" & FRectMakerid & "' " + VbCrlf
		sqlStr = sqlStr + " 	and d.upcheorderlinkcode is null" + VbCrlf
		sqlStr = sqlStr + " 	and m.deldt is null " + vbCrlf
		sqlStr = sqlStr + " 	and d.deldt is null " + vbCrlf
		'response.write sqlStr + "<br><br>"
		rsget.Open sqlStr, dbget, 1

		'오프 or 업체배송 상품만(온라인상품 주문은 온라인상품주문에 합쳐서 별도로 한다.)
		'주문접수/주문확인 상태의 주문에서 상품정보를 가져온다.
		sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail( " + vbCrlf
		sqlStr = sqlStr + " 	masteridx " + vbCrlf
		sqlStr = sqlStr + " 	, itemgubun " + vbCrlf
		sqlStr = sqlStr + " 	, makerid " + vbCrlf
		sqlStr = sqlStr + " 	, itemid " + vbCrlf
		sqlStr = sqlStr + " 	, itemoption " + vbCrlf
		sqlStr = sqlStr + " 	, itemname " + vbCrlf
		sqlStr = sqlStr + " 	, itemoptionname " + vbCrlf
		sqlStr = sqlStr + " 	, sellcash " + vbCrlf
		sqlStr = sqlStr + " 	, suplycash " + vbCrlf
		sqlStr = sqlStr + " 	, buycash " + vbCrlf
		sqlStr = sqlStr + " 	, baljuitemno " + vbCrlf
		sqlStr = sqlStr + " 	, realitemno " + vbCrlf
		sqlStr = sqlStr + " 	, baljudiv " + vbCrlf
		sqlStr = sqlStr + " ) " + vbCrlf
		sqlStr = sqlStr + " select  " + vbCrlf
		sqlStr = sqlStr + " 	" & CStr(iid) & " " + vbCrlf
		sqlStr = sqlStr + " 	, d.itemgubun " + vbCrlf
		sqlStr = sqlStr + " 	, MAX(d.makerid) " + vbCrlf
		sqlStr = sqlStr + " 	, d.itemid " + vbCrlf
		sqlStr = sqlStr + " 	, d.itemoption " + vbCrlf
		sqlStr = sqlStr + " 	, MAX(d.itemname) " + vbCrlf
		sqlStr = sqlStr + " 	, MAX(d.itemoptionname) " + vbCrlf
		sqlStr = sqlStr + " 	, MAX(d.sellcash) " + vbCrlf
		sqlStr = sqlStr + " 	, MAX(d.suplycash) as suplycash " + vbCrlf   '' 공급가가 다른 CASE 가 있음. group by 로 변경 2016/05/24
		sqlStr = sqlStr + " 	, MIN(d.buycash) " + vbCrlf
		sqlStr = sqlStr + " 	, 0 " + vbCrlf				'수량은 아래에서 입력한다.
		sqlStr = sqlStr + " 	, 0 " + vbCrlf
		sqlStr = sqlStr + " 	, '0' " + vbCrlf
		sqlStr = sqlStr + " from " + vbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m " + vbCrlf
		sqlStr = sqlStr + " 	, [db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
		sqlStr = sqlStr + " where " + vbCrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
		sqlStr = sqlStr + " 	and m.idx=d.masteridx " + vbCrlf
		sqlStr = sqlStr + " 	and m.statecd in ('0', '1') " + vbCrlf
		sqlStr = sqlStr + " 	and d.makerid='" & FRectMakerid & "' " + vbCrlf
		sqlStr = sqlStr + " 	and m.baljucode in (" + CStr(FRectOrgBaljuCode) + ") " + vbCrlf
		sqlStr = sqlStr + " 	and m.deldt is null " + vbCrlf
		sqlStr = sqlStr + " 	and d.deldt is null " + vbCrlf
		''sqlStr = sqlStr + " group by d.itemgubun,d.makerid, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.sellcash, d.buycash"
		sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption"		'// 하나로 합친다., skyer9, 2018-02-19

'		sqlStr = sqlStr + " select distinct " + vbCrlf
'		sqlStr = sqlStr + " 	" & CStr(iid) & " " + vbCrlf
'		sqlStr = sqlStr + " 	, d.itemgubun " + vbCrlf
'		sqlStr = sqlStr + " 	, d.makerid " + vbCrlf
'		sqlStr = sqlStr + " 	, d.itemid " + vbCrlf
'		sqlStr = sqlStr + " 	, d.itemoption " + vbCrlf
'		sqlStr = sqlStr + " 	, d.itemname " + vbCrlf
'		sqlStr = sqlStr + " 	, d.itemoptionname " + vbCrlf
'		sqlStr = sqlStr + " 	, d.sellcash " + vbCrlf
'		sqlStr = sqlStr + " 	, d.suplycash " + vbCrlf
'		sqlStr = sqlStr + " 	, d.buycash " + vbCrlf
'		sqlStr = sqlStr + " 	, 0 " + vbCrlf				'수량은 아래에서 입력한다.
'		sqlStr = sqlStr + " 	, 0 " + vbCrlf
'		sqlStr = sqlStr + " 	, '0' " + vbCrlf
'		sqlStr = sqlStr + " from " + vbCrlf
'		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m " + vbCrlf
'		sqlStr = sqlStr + " 	, [db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
'		sqlStr = sqlStr + " where " + vbCrlf
'		sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
'		sqlStr = sqlStr + " 	and m.idx=d.masteridx " + vbCrlf
'		sqlStr = sqlStr + " 	and m.statecd in ('0', '1') " + vbCrlf
'		sqlStr = sqlStr + " 	and d.makerid='" & FRectMakerid & "' " + vbCrlf
'		sqlStr = sqlStr + " 	and m.baljucode in (" + CStr(FRectOrgBaljuCode) + ") " + vbCrlf
'		sqlStr = sqlStr + " 	and m.deldt is null " + vbCrlf
'		sqlStr = sqlStr + " 	and d.deldt is null " + vbCrlf
		'response.write sqlStr + "<br><br>"

		rsget.Open sqlStr, dbget, 1

		itemgubunarr 	= split(FRectITemGubunArr,",")
		itemidarr 		= split(FRectItemIdArr,",")
		itemoptionarr 	= split(FRectItemOptionArr,",")
		ordernoarr 		= split(FRectOrderNoArr,",")

		cnt = ubound(itemgubunarr)

		for i = 0 to cnt

			if (Trim(itemgubunarr(i)) <> "") then

				sqlStr = " update d " + vbCrlf
				sqlStr = sqlStr + " set " + vbCrlf
				sqlStr = sqlStr + " 	d.baljuitemno = d.baljuitemno + " & Trim(ordernoarr(i)) & " " + vbCrlf
				sqlStr = sqlStr + " 	, d.realitemno = d.realitemno + " & Trim(ordernoarr(i)) & " " + vbCrlf
				sqlStr = sqlStr + " from " + vbCrlf
				sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
				sqlStr = sqlStr + " where " + vbCrlf
				sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
				sqlStr = sqlStr + " 	and d.masteridx = " & CStr(iid) & " " + vbCrlf
				''sqlStr = sqlStr + " 	and d.makerid='" & FRectMakerid & "' " + vbCrlf
				sqlStr = sqlStr + " 	and d.itemgubun='" & Trim(itemgubunarr(i)) & "' " + vbCrlf
				sqlStr = sqlStr + " 	and d.itemid=" & Trim(itemidarr(i)) & " " + vbCrlf
				sqlStr = sqlStr + " 	and d.itemoption='" & Trim(itemoptionarr(i)) & "' " + vbCrlf
				'response.write sqlStr + "<br><br>"

				rsget.Open sqlStr, dbget, 1

			end if

		next

		'수량이 0 이면 삭제
		sqlStr = " delete " + vbCrlf
		sqlStr = sqlStr + " 	d " + vbCrlf
		sqlStr = sqlStr + " from " + vbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
		sqlStr = sqlStr + " where " + vbCrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
		sqlStr = sqlStr + " 	and d.masteridx = " & CStr(iid) & " " + vbCrlf
		sqlStr = sqlStr + " 	and d.baljuitemno = 0 " + vbCrlf
		sqlStr = sqlStr + " 	and d.makerid='" & FRectMakerid & "' " + vbCrlf
		rsget.Open sqlStr, dbget, 1

		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
		sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
		sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T" + vbCrlf
		sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)

		rsget.Open sqlStr, dbget, 1

		MakeUpcheJumunNew = iid
	end function

	public Sub GetOneOrderSheetMaster_IMSI()
		dim sqlStr

		sqlStr = " select top 1"
		sqlStr = sqlStr + " b.* "
		sqlStr = sqlStr + " from [db_storage].[dbo].[tbl_ordersheet_master_before_save] b "
		sqlStr = sqlStr + " where b.idx=" + CStr(FRectIdx) + ""

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		set FOneItem = new COrderSheetMasterItem
		if Not rsget.Eof then
			FOneItem.Fidx            = rsget("idx")
			FOneItem.Ftargetid       = rsget("targetid")
			FOneItem.Fbaljuid        = rsget("baljuid")
			FOneItem.Fregdate        = rsget("regdate")
			FOneItem.Fupdt           = rsget("updt")
			FOneItem.Fdeldt          = rsget("deldt")
			FOneItem.Fscheduledate   = rsget("scheduledate")
			FOneItem.Fsongjangdiv    = rsget("songjangdiv")
			FOneItem.Fsongjangno     = rsget("songjangno")
			FOneItem.Fcomment        = db2html(rsget("comment"))
		end if
		rsget.Close

	end Sub

	public Sub GetOrderSheetDetail_IMSI()
		dim i,sqlStr


		sqlStr = " SELECT d.* "
		sqlStr = sqlStr + " 	,IsNull(si.shopitemname, i.itemname) AS itemname "
		sqlStr = sqlStr + " 	,IsNull(si.shopitemoptionname, IsNull(o.optionname, '')) AS itemoptionname "
		sqlStr = sqlStr + " 	,IsNull(si.makerid, i.makerid) AS makerid "
		sqlStr = sqlStr + " FROM "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail_before_save d "
		sqlStr = sqlStr + " 	LEFT JOIN [db_shop].[dbo].tbl_shop_item si "
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun <> '10' "
		sqlStr = sqlStr + " 		AND d.itemgubun = si.itemgubun "
		sqlStr = sqlStr + " 		AND d.itemid = si.shopitemid "
		sqlStr = sqlStr + " 		AND d.itemoption = si.itemoption "
		sqlStr = sqlStr + " 	LEFT JOIN [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun = '10' "
		sqlStr = sqlStr + " 		AND d.itemid = i.itemid "
		sqlStr = sqlStr + " 	LEFT JOIN [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun = '10' "
		sqlStr = sqlStr + " 		AND d.itemid = i.itemid "
		sqlStr = sqlStr + " 		AND o.itemid = i.itemid "
		sqlStr = sqlStr + " 		AND d.itemoption = o.itemoption "
		sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " 	AND d.deldt IS NULL "
		sqlStr = sqlStr + " order by IsNull(si.makerid, i.makerid), d.itemgubun, d.itemid, d.itemoption"

		''response.write sqlStr & "<br>"
		''response.end
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).Fidx            	= rsget("idx")
				FItemList(i).Fmasteridx      	= rsget("masteridx")
				FItemList(i).Fitemgubun      	= rsget("itemgubun")
				FItemList(i).Fitemid         	= rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fbaljuitemno   	= rsget("itemno")

				FItemList(i).Fsellcash       	= rsget("sellcash")
				FItemList(i).Fsuplycash      	= rsget("suplycash")
				FItemList(i).Fbuycash       	= rsget("buycash")

				FItemList(i).FMakerid			= rsget("makerid")
				FItemList(i).Fitemname       	= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname 	= db2html(rsget("itemoptionname"))

				'// idx, masteridx, itemgubun, itemid, itemoption, sellcash, suplycash, buycash, itemno, regdate, updt, deldt

				''''FItemList(i).Fitemname       = db2html(rsget("itemname"))
				''				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				''
				''
				''
				''				FItemList(i).Fbaljuitemno   	= rsget("baljuitemno")				'// 주문수량
				''				FItemList(i).Frealbaljuitemno   = rsget("realbaljuitemno")			'// 발주수량
				''				FItemList(i).Frealitemno    	= rsget("realitemno")				'// 확정수량
				''				FItemList(i).Fcheckitemno    	= rsget("checkitemno")				'// 검품수량
				''				FItemList(i).Fregdate       = rsget("regdate")
				''				FItemList(i).Fupdt         = rsget("updt")
				''				FItemList(i).Fdeldt        = rsget("deldt")
				''				FItemList(i).Fbaljudiv     = rsget("baljudiv")
				''				FItemList(i).Fcomment      = db2html(rsget("comment"))
				''				FItemList(i).FMakerid	= rsget("makerid")
				''				FItemList(i).FItemDefaultMwDiv	= rsget("mwdiv")
				''				FItemList(i).Fdeliverytype = rsget("deliverytype")
				''				FItemList(i).Fonlinesellcash = rsget("onlinesellcash")
				''				FItemList(i).Fonlinebuycash  = rsget("onlinebuycash")
				''				FItemList(i).Fshopdefaultmargin = rsget("shopdefaultmargin")
				''				FItemList(i).Fshopdefaultsuplymargin = rsget("shopdefaultsuplymargin")
				''				FItemList(i).FoffChargeDiv = rsget("offchargediv")
				''				FItemList(i).Fipgoflag = rsget("ipgoflag")
				''				FItemList(i).Fdefaultmaginflag = rsget("defaultmaginflag")
				''				FItemList(i).Fbuymaginflag = rsget("buymaginflag")
				''				FItemList(i).Fsuplymaginflag = rsget("suplymaginflag")
				''				FItemList(i).FPublicBarcode = rsget("barcode")
				''
				''				FItemList(i).Fsmallimage = rsget("smallimage")
				''				FItemList(i).Foffimgsmall = rsget("offimgsmall")
				''				if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
				''				if isnull(FItemList(i).Foffimgsmall) then FItemList(i).Foffimgsmall=""
				''				if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage
				''				if FItemList(i).Foffimgsmall<>"" then FItemList(i).Foffimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgsmall
				''
				''				FItemList(i).Fbasicimage = rsget("basicimage")
				''				FItemList(i).Foffimgmain = rsget("offimgmain")
				''				if isnull(FItemList(i).Fbasicimage) then FItemList(i).Fbasicimage=""
				''				if isnull(FItemList(i).Foffimgmain) then FItemList(i).Foffimgmain=""
				''				if FItemList(i).Fbasicimage<>"" then FItemList(i).Fbasicimage     = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fbasicimage
				''				if FItemList(i).Foffimgmain<>"" then FItemList(i).Foffimgmain = "http://webimage.10x10.co.kr/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Foffimgmain
				''
				''				FItemList(i).Fdetail_status = rsget("detail_status")
				''				FItemList(i).Fdetail_description = rsget("detail_description")
				''				FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
				''				FItemList(i).Fboxsongjangno  = rsget("boxsongjangno")
				''				FItemList(i).FUpcheManageCode = rsget("upchemanagecode")
				''				FItemList(i).Forgsellprice = rsget("orgsellprice")
				''
				''				if isnull(FItemList(i).Fcheckitemno) then FItemList(i).Fcheckitemno = 0

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		total_jumunsellcash=0
		total_jumunsuplycash=0
		total_totalsuplycash=0
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


public Fcompanyno
public FpurchaseType

 	public Function fnGetBrandInfo
 		dim strSql
 		strSql = "select company_no,purchaseType   from   db_partner.dbo.tbl_partner   where id ='"&FRectMakerid&"' "
		rsget.Open strSql, dbget, 1
		if not rsget.eof then
			Fcompanyno		= rsget("company_no")
			FpurchaseType = rsget("purchaseType")
	end if
	rsget.close
	end Function

end Class

Function fnGetAGVCheckBalju(baljucode)
	dim sqlStr
	sqlStr = "select top 1 idx from [db_aLogistics].[dbo].tbl_agv_scheduleditems where requestMaster='STOCKIN("&baljucode&")' and isusing='Y'"
	rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	if not rsget_Logistics.eof then
		fnGetAGVCheckBalju = True
	else
		fnGetAGVCheckBalju = False
	end if
	rsget_Logistics.close
end Function

'// 기준일 처리 	'/2012.06.01 한용민 생성
function drawipgo_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>주문일</option>
	<option value='ipgo' <%if selectedId="ipgo" then response.write " selected"%>>출고일</option>
</select>
<%
end function

'//물류주문 상태 셀렉트박스		'/2013.06.19 한용민 생성
function drawstatecd(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>CHOICE</option>
	<option value="foreign0" <% if selectedId="foreign0" then response.write "selected" %> >업체접수(견적요청)</option>
	<option value="foreign3" <% if selectedId="foreign3" then response.write "selected" %> >업체접수확인</option>
	<option value="foreign7" <% if selectedId="foreign7" then response.write "selected" %> >업체접수완료</option>
	<option value=" " <% if selectedId=" " then response.write "selected" %> >주문서작성중</option>
	<option value="0" <% if selectedId="0" then response.write "selected" %> >주문접수</option>
	<option value="before1" <% if selectedId="before1" then response.write "selected" %> >주문확인 이전전체</option>
	<option value="1" <% if selectedId="1" then response.write "selected" %> >주문확인</option>
	<option value="2" <% if selectedId="2" then response.write "selected" %> >입금대기</option>
	<option value="5" <% if selectedId="5" then response.write "selected" %> >배송준비</option>
	<option value="6" <% if selectedId="6" then response.write "selected" %> >출고대기</option>
	<option value="7" <% if selectedId="7" then response.write "selected" %> >출고완료</option>
	<option value="8" <% if selectedId="8" then response.write "selected" %> >입고대기</option>
	<option value="9" <% if selectedId="9" then response.write "selected" %> >입고완료</option>
</select>
<%
end function

function GetPurchaseTypeList(codeList, ByRef makerid, ByRef purchasetypestr, ByRef formidx)
    dim sqlStr, str

	sqlStr = " select distinct top 1 p.purchasetype, p.id as makerid "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_ordersheet_master] om "
    sqlStr = sqlStr & " 	join db_partner.dbo.tbl_partner p on om.targetid = p.id "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and om.baljucode in ('" & Replace(codeList, ",", "','") & "') "
    ''response.write sqlStr

    makerid = ""
    purchasetypestr = ""
    formidx = ""

    rsget.Open sqlStr, dbget, 1
    if Not rsget.eof then
        select case rsget("purchasetype")
            case "1", "4", "5":
                formidx = 102
                purchasetypestr = "상품사입"
            case "6", "7", "9":
                formidx = 103
                purchasetypestr = "상품수입"
            case "8", "3":
                formidx = 104
                purchasetypestr = "상품제작"
            case else
                '
        end select
        makerid = rsget("makerid")
    end if
	rsget.close

end function

function ordersheetsmssend(masteridx)
	dim sqlstr

	if masteridx="" or isnull(masteridx) then exit function

	sqlstr = "update db_storage.dbo.tbl_ordersheet_master set smssenddate=getdate() where idx="& masteridx &""

	dbget.execute sqlstr
end function
%>
