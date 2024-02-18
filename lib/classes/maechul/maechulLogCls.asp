<%
'###########################################################
' Description : 매출로그
' Hieditor : 2011.11.01 이상구 생성
'			 2013.11.14 한용민 수정
'###########################################################

Class CMaechulLogItem
	public Fyyyymm
	public Fyyyymm2
	public fpurchasetypename
	public FtargetGbn
	public Forderserial
	public Fsuborderserial

	public Forgorderserial

	public FactDivCode
	public FactDate
	public Fjumundiv
	public Faccountdiv
	public fmwdiv_beasongdiv
	public Fitemid
	public Fitemoption
	public Fitemno
    public Fsku
	public Fitemname
	public Fitemoptionname
	public Fmakerid
	public Fvatinclude
	public Fomwdiv
	public FanbunPriceDetailSUM
	public ForgTotalPrice							'// 소비자가
	public FsubtotalpriceCouponNotApplied			'// 판매가(할인가)
	public Ftotalsum								'// 상품쿠폰적용가
	public FtotalReducedPrice						'// 보너스쿠폰적용가
	public FtotalReducedPriceBeasongPay				'// 보너스쿠폰적용가(배송비만)
	public FtotalItemCouponDiscount					'// 상품쿠폰할인액
	public FtotalBonusCouponDiscount				'// 보너스쿠폰할인액
	public FtotalBeasongBonusCouponDiscount			'// 보너스쿠폰할인액(배송비쿠폰만)
	public FtotalPriceBonusCouponDiscount			'// 보너스쿠폰할인액(금액쿠폰액만)
	public Fallatdiscountprice						'// 기타할인
	public FtotalMaechulPrice						'// 매출총액
	public FtotalMaechulVatPrice					'// 부가세
	public FmileTotalPrice							'// 마일리지사용액
	public FgiftTotalPrice							'// 기프트카드사용액
	public FdepositTotalPrice						'// 예치금사용액
	public FtotalBuycash							'// 매입가
	public FtotalBuycashVAT
	public FtotalBuycashCouponNotApplied
	public FtotalUpcheJungsanCash					'// 업체정산액
	public FtotalUpcheJungsanCashVAT
	public FtotalMileage							'// 부여마일리지
	public Fbeasongdate
	public FDTLjFixedDt								'// 정산확정일 20200130
	public Fcanceldate
	public Fipkumdate
	public Fsitename
	public Frdsite
	public Fregdate
	public ftotalMaechulPrice_M
	public ftotalMaechulPrice_W
	public ftotalMaechulPrice_U
	public ftotalMaechulPrice_TT
	public ftotalMaechulPrice_UU
	public faccountMaechulPrice_M
	public faccountMaechulPrice_W
	public faccountMaechulPrice_U
	public ftotalUpcheJungsanCash_M
	public ftotalUpcheJungsanCash_W
	public ftotalUpcheJungsanCash_U
	public fbeasongUpcheJungsanCash_TT
	public fbeasongUpcheJungsanCash_UU
	public FrealTotalsum
	public FrealSpendmileage
	public FrealGainmileage

	public FavgipgoPrice
	public FoverValueStockPrice

	Public Fpurchasetype
	Public Fcatename
	Public FChannelName
	Public Fbeadaldiv

	public Function getMeaChulGubunName()
		getMeaChulGubunName = ""

		if (fmwdiv_beasongdiv="M" or fmwdiv_beasongdiv="B031") then
			getMeaChulGubunName = GetVatIncludeName
		elseif (fmwdiv_beasongdiv="U" or fmwdiv_beasongdiv="W" or fmwdiv_beasongdiv="B013" or fmwdiv_beasongdiv="B012") then
			getMeaChulGubunName = "과세"

		elseif (fmwdiv_beasongdiv="TT" or fmwdiv_beasongdiv="UU" or fmwdiv_beasongdiv="PP" or fmwdiv_beasongdiv="R") then
			getMeaChulGubunName = "과세"
		end if
	end function

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "일반유통"
    	ELSEIF FPurchasetype = "3" then
    	    getPurchasetypeName = "PB"
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

	public function GetFullOrderSerial
		GetFullOrderSerial = Forderserial + "-" + Format00(3, Fsuborderserial)
	end function

	public function GetRealPayPrice
		GetRealPayPrice = FtotalMaechulPrice - (FmileTotalPrice + FgiftTotalPrice + FdepositTotalPrice)
	end function

	public function GetActDivCodeName
		Select Case FactDivCode
			Case "A"
				GetActDivCodeName = "원주문"
			Case "C"
				GetActDivCodeName = "취소주문"
			Case "H"
				GetActDivCodeName = "상품변경"
			Case "E"
				GetActDivCodeName = "교환주문"
			Case "M"
				GetActDivCodeName = "반품주문"
			Case "CC"
				GetActDivCodeName = "취소정상화"
			Case "HH"
				GetActDivCodeName = "상품변경취소"
			Case "EE"
				GetActDivCodeName = "교환취소"
			Case "MM"
				GetActDivCodeName = "반품취소"
			Case Else
				GetActDivCodeName = GetActDivCodeName
		End Select
	end function

	public function GetVatIncludeName
		Select Case Fvatinclude
			Case "N"
				GetVatIncludeName = "면세"
			Case Else
				GetVatIncludeName = "과세"
		End Select
	end function

	public function GetOMWdivName
		if (Fitemid = 0) then
'			if (Left(Fitemoption, 1) = "9") then
'				GetOMWdivName = "업배"
'			else
'				GetOMWdivName = "텐배"
'			end if
            if (Fomwdiv="UU") then
                GetOMWdivName = "업배"
            elseif (Fomwdiv="TT") then
                GetOMWdivName = "텐배"
            else
                GetOMWdivName = Fomwdiv
            end if
		else
			Select Case Fomwdiv
				Case "M"
					GetOMWdivName = "매입"
				Case "W"
					GetOMWdivName = "위탁"
				Case "U"
					GetOMWdivName = "업체"
				Case "R"
					GetOMWdivName = "랜탈"

				Case "B000"
					GetOMWdivName = "미지정"
				Case "B011"
					GetOMWdivName = "위탁판매"
				Case "B012"
					GetOMWdivName = "업체위탁"
				Case "B013"
					GetOMWdivName = "출고위탁"
				Case "B021"
					GetOMWdivName = "오프매입"
				Case "B022"
					GetOMWdivName = "매장매입"
				Case "B023"
					GetOMWdivName = "가맹점매입"
				Case "B031"
					GetOMWdivName = "출고매입"
				Case "B032"
					GetOMWdivName = "센터매입"
				Case "B999"
					GetOMWdivName = "기타보정"
				Case "PP"
					GetOMWdivName = "포장"
				Case Else
					GetOMWdivName = Fomwdiv
			End Select
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="14" then
			JumunMethodName="편의점결제"
		elseif Faccountdiv="100" then
			JumunMethodName="신용"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="입점몰"
		elseif Faccountdiv="80" then
			JumunMethodName="All@"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+신용"
		elseif Faccountdiv="400" then
			JumunMethodName="핸드폰"
		elseif Faccountdiv="550" then
			JumunMethodName="기프팅"
		elseif Faccountdiv="560" then
			JumunMethodName="기프티콘"
		end if
	end Function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CMaechulLogSUMItem
	public ForgTotalPrice							'// 소비자가
	public FsubtotalpriceCouponNotApplied			'// 판매가(할인가)
	public Ftotalsum								'// 상품쿠폰적용가
	public FtotalReducedPrice						'// 보너스쿠폰적용가
	public FtotalReducedPriceBeasongPay				'// 보너스쿠폰적용가(배송비만)
	public FtotalItemCouponDiscount					'// 상품쿠폰할인액
	public FtotalBonusCouponDiscount				'// 보너스쿠폰할인액
	public FtotalBeasongBonusCouponDiscount			'// 보너스쿠폰할인액(배송비쿠폰만)
	public FtotalPriceBonusCouponDiscount			'// 보너스쿠폰할인액(금액쿠폰액만)
	public Fallatdiscountprice						'// 기타할인
	public FtotalMaechulPrice						'// 매출총액
	public FtotalMaechulVatPrice					'// 부가세
	public FtotalBuycash							'// 매입가
	public FtotalBuycashVAT
	public FtotalBuycashCouponNotApplied
	public FtotalUpcheJungsanCash					'// 업체정산액
	public FtotalUpcheJungsanCashVAT
	public FtotalMileage							'// 부여마일리지
	public FavgipgoPrice
	public FmileTotalPrice
	public FgiftTotalPrice
	public FdepositTotalPrice

	Public Fyyyymm
	Public Fsitename
	Public FsellChannelName

	Public ForgOrderCnt
	Public FcancelOrderCnt
	Public FreturnOrderCnt

    public Fuserlevel

	public function GetRealPayPrice
		GetRealPayPrice = FtotalMaechulPrice - (FmileTotalPrice + FgiftTotalPrice + FdepositTotalPrice)
	end Function

	public function GetSellChannelName
		GetSellChannelName = FsellChannelName
		If (FsellChannelName <> "???") Then
			Exit Function
		End If

		Select Case Left(Fsitename,3)
			Case "aca"
				GetSellChannelName = "ACA"
			Case "diy"
				GetSellChannelName = "DIY"
			Case "str"
				''GetSellChannelName = Right(Fsitename,3)
                GetSellChannelName = "OFF"
			Case Else
				'//
		End Select
	end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CMaechulPaymentLogCheckItem
	public FtargetGbn
	public Factdate
	public Forderserial
	public FtotalOrderMaechulPrice
	public FtotalpayreqPrice

    public FappPrice
    public FrealPayPrice

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CMaechulPaymentLogItem
	public FtargetGbn
	public Forderserial
	public Fsuborderserial
	public Forgorderserial
	public Fchgorderserial
	public FpayDivCode
	public FPGgubun
	public FPGuserid
	public FPGkey
	public FPGCSkey
	public FpayReqPrice
	public FrealPayPrice
	public FpayReqDate
	public FpayDate
	public FmaeipDate
	public FmayIpkumDate
	public FactDivCode
	public Fregdate

	public FmayPayMethod

	public FmatchMethod

    public FcommPrice       ''2014/01/07 추가
    public FcommVatPrice
    public FjungsanPrice    ''2014/01/07 추가

	public function GetFullOrderSerial
		GetFullOrderSerial = Forderserial + "-" + Format00(3, Fsuborderserial)
	end function

	public function GetMatchMethodName
		Select Case FmatchMethod
			Case "X"
				GetMatchMethodName = "매칭이전"
			Case "A"
				GetMatchMethodName = "자동매칭"
			Case "H"
				GetMatchMethodName = "수기매칭"
			Case "R"
				GetMatchMethodName = "환불진행중"
			Case Else
				GetMatchMethodName = FmatchMethod
		End Select
	end function

	public function GetActDivCodeName
		Select Case FactDivCode
			Case "A"
				GetActDivCodeName = "원주문"
			Case "C"
				GetActDivCodeName = "취소주문"
			Case "H"
				GetActDivCodeName = "상품변경"
			Case "E"
				GetActDivCodeName = "교환주문"
			Case "M"
				GetActDivCodeName = "반품주문"
			Case "CC"
				GetActDivCodeName = "취소정상화"
			Case "HH"
				GetActDivCodeName = "상품변경취소"
			Case "EE"
				GetActDivCodeName = "교환취소"
			Case "MM"
				GetActDivCodeName = "반품취소"
			Case Else
				GetActDivCodeName = FactDivCode
		End Select
	end function

	Public function GetPayDivCodeName()
		if FpayDivCode = "7" then
			GetPayDivCodeName = "무통장"
		elseif FpayDivCode = "100" then
			GetPayDivCodeName = "신용"
		elseif FpayDivCode = "20" then
			GetPayDivCodeName = "실시간"
		elseif FpayDivCode = "30" then
			GetPayDivCodeName = "포인트"
		elseif FpayDivCode = "50" then
			GetPayDivCodeName = "입점몰"
		elseif FpayDivCode = "80" then
			GetPayDivCodeName = "All@"
		elseif FpayDivCode = "90" then
			GetPayDivCodeName = "상품권"
		elseif FpayDivCode = "110" then
			GetPayDivCodeName = "OK캐시백"
		elseif FpayDivCode = "400" then
			GetPayDivCodeName = "핸드폰"
		elseif FpayDivCode = "550" then
			GetPayDivCodeName = "기프팅"
		elseif FpayDivCode = "560" then
			GetPayDivCodeName = "기프티콘"
		elseif FpayDivCode = "mil" then
			GetPayDivCodeName = "마일리지"
		elseif FpayDivCode = "dep" then
			GetPayDivCodeName = "예치금"
		elseif FpayDivCode = "gif" then
			GetPayDivCodeName = "기프트카드"
		elseif FpayDivCode = "77" then
			GetPayDivCodeName = "무통장환불"
		elseif FpayDivCode = "6" then
			GetPayDivCodeName = "무통장입금"
		elseif FpayDivCode = "rmi" then
			GetPayDivCodeName = "마일리지환불"
		elseif FpayDivCode = "rde" then
			GetPayDivCodeName = "예치금환불"
		elseif FpayDivCode = "0" then
			GetPayDivCodeName = "결제없음"
		else
			GetPayDivCodeName = FpayDivCode
		end if
	end function

	Public function GetMayPayMethodName()
		if FmayPayMethod = "7" then
			GetMayPayMethodName = "무통장"
		elseif FmayPayMethod = "100" then
			GetMayPayMethodName = "신용"
		elseif FmayPayMethod = "20" then
			GetMayPayMethodName = "실시간"
		elseif FmayPayMethod = "30" then
			GetMayPayMethodName = "포인트"
		elseif FmayPayMethod = "50" then
			GetMayPayMethodName = "입점몰"
		elseif FmayPayMethod = "80" then
			GetMayPayMethodName = "All@"
		elseif FmayPayMethod = "90" then
			GetMayPayMethodName = "상품권"
		elseif FmayPayMethod = "110" then
			GetMayPayMethodName = "OK+신용"
		elseif FmayPayMethod = "400" then
			GetMayPayMethodName = "핸드폰"
		elseif FmayPayMethod = "550" then
			GetMayPayMethodName = "기프팅"
		elseif FmayPayMethod = "560" then
			GetMayPayMethodName = "기프티콘"
		elseif FmayPayMethod = "mil" then
			GetMayPayMethodName = "마일리지"
		elseif FmayPayMethod = "dep" then
			GetMayPayMethodName = "예치금"
		elseif FmayPayMethod = "gif" then
			GetMayPayMethodName = "기프트카드"
		elseif FmayPayMethod = "77" then
			GetMayPayMethodName = "무통장환불"
		elseif FmayPayMethod = "rmi" then
			GetMayPayMethodName = "마일리지환불"
		elseif FmayPayMethod = "rde" then
			GetMayPayMethodName = "예치금환불"
		else
			GetMayPayMethodName = FmayPayMethod
		end if
	end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CMaechulLog
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FPageCount

	public tendb
	public FRectTargetGbn
	public FRectmakerid
	public FRectActDivCode
    public FRectDategbn
	public FRectStartDate
	public FRectEndDate
	public FRectStartDate2
	public FRectEndDate2
	public FRectOrgPayStartDate
	public FRectOrgPayEndDate
	public FRectActDateStartDate
	public FRectActDateEndDate
	public FRectChulgoDateStartDate
	public FRectChulgoDateEndDate
	public FRectjFixedDtStartDate
	public FRectjFixedDtEndDate
	public FRectSearchField
	public FRectSearchText
	public FRectPayDivCode
	public FRectmwdiv_beasongdiv
	public FRectMichulgoOnly
    public FRectMiJungsanOnly
	public FRectvatinclude
	public FRectChkGrpByOrderserial
	public FRectChkOnlyDiff
	public FRectShowOnlyPriceNotMatch
	public FRectMatchState
    public FRectSitename
    public FRectIncActPayMonthDiff

	public FRectExcNoPay
	public FRectExcNoReqPay
	public FRectExcHP
	Public FRectExcGift
    public FRectExceptDlv  ''배송비제외

    public FRectAddDategbn
    public FRectAddStartDate
    public FRectAddEndDate
    public FRectExceptSite  ''해당매출처제외

	public FRectExcTPL
	Public FRectUseNewDB
	public FRectShowStatistic
	public FRectshowOnlyStatistic
	public FRectExcZeroPrice

	public FRectPGgubun
	public FRectPGuserid

	public FRectDateGubun

	public FRectExc6month

	public FRectGrpBy

	Public FRectPurchaseType
	public FRectNextMonthJungsanFixed

    public FTotalPayReqPrice
    public FTotalRealPayPrice

    public FRectShowLevel

	'/admin/maechul/maechul_month_sitename_log.asp
	public function GetMaechul_month_sitename_Log()
		dim sqlStr,i, groupsqlStr, fieldsqlStr, addSqlStr, indexmSqlStr, indexdSqlStr


		if FRectStartDate="" or FRectEndDate="" then exit function
''주석처리 2015/04/02
		if FRectDategbn="ActDate" then
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_actDate))"
		elseif FRectDategbn="chulgoDate" then
			indexdSqlStr = indexdSqlStr + " with (index(IX_tbl_order_detail_log_beasongdate))"
		else
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_ipkumdate))"
		end if

		if FRectDategbn="ActDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.actDate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.actDate,21) as yyyymm"
		elseif FRectDategbn="chulgoDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),d.beasongdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),d.beasongdate,21) as yyyymm"
		else
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.ipkumdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.ipkumdate,21) as yyyymm"
		end if

'		''2차검색날짜
'		if (FRectAddDategbn<>"") then
'    		if FRectAddDategbn="ActDate" then
'    	        if FRectAddStartDate <> "" then
'    				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectAddStartDate) + "'"
'    			end if
'    			if FRectAddEndDate <> "" then
'    				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectAddEndDate) + "'"
'    			end if
'
'    		elseif FRectDategbn="chulgoDate" then
'    	        if FRectAddStartDate <> "" then
'    				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectAddStartDate) + "'"
'    			end if
'    			if FRectAddEndDate <> "" then
'    				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectAddEndDate) + "'"
'    			end if
'
'    		else
'    	        if FRectAddStartDate <> "" then
'    				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectAddStartDate) + "'"
'    			end if
'    			if FRectAddEndDate <> "" then
'    				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectAddEndDate) + "'"
'    			end if
'
'    		end if
'		end if

		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if FRectmwdiv_beasongdiv="M" or FRectmwdiv_beasongdiv="W" or FRectmwdiv_beasongdiv="U" or FRectmwdiv_beasongdiv="R" then
			addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + FRectmwdiv_beasongdiv + "'"
		elseif FRectmwdiv_beasongdiv="TT" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
		elseif FRectmwdiv_beasongdiv="UU" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
		end if
		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if
		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "'"
		end if

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	m.sitename " & fieldsqlStr		'/출고처
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by m.sitename " & groupsqlStr
		sqlStr = sqlStr & " ) as t"

'		response.write sqlStr &"<br>"
'		response.end
'		db3_rsget.Open sqlStr,db3_dbget,1
'			FTotalCount = db3_rsget("cnt")
'		db3_rsget.Close

'		if FTotalCount<1 then exit function

		sqlStr = "select *"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " 	m.sitename " & fieldsqlStr		'/출고처
		sqlStr = sqlStr & " 	, IsNull(Sum(d.orgitemcost*d.itemno), 0) as orgTotalPrice"  	'/소비자가
		sqlStr = sqlStr & " 	, IsNull(Sum(d.itemcostCouponNotApplied * d.itemno), 0) as subtotalpriceCouponNotApplied"		'/판매가(할인가)
		sqlStr = sqlStr & " 	, IsNull(Sum(d.itemcost*d.itemno), 0) as totalsum"		'/상품쿠폰적용가
		sqlStr = sqlStr & " 	, IsNull(sum((d.itemcost-d.reducedPrice-IsNull(d.allAtDiscount, 0))*d.itemno), 0) as totalBonusCouponDiscount"
		sqlStr = sqlStr & " 	, IsNull(sum((case when d.itemid=0 then d.itemcost-d.reducedPrice else 0 end)*d.itemno), 0) as totalBeasongBonusCouponDiscount"		'/배송비쿠폰
		sqlStr = sqlStr & " 	, IsNull(sum(d.anbunCouponPriceDetailSUM), 0) as totalPriceBonusCouponDiscount"		'/정액쿠폰
		sqlStr = sqlStr & " 	, IsNull(sum(IsNull(d.allAtDiscount, 0)*d.itemno), 0) as allatdiscountprice" 		'/기타할인(올앳)
	sqlStr = sqlStr & " 	, IsNull(sum(d.anbunAppliedPriceDetailSUM), 0) as totalMaechulPrice"		'/매출총액
		sqlStr = sqlStr & " 	, IsNull(sum(d.upcheJungsanCash*d.itemno), 0) as totalUpcheJungsanCash"		'/업체정산액
		sqlStr = sqlStr & " 	, IsNull(sum(d.mileage * d.itemno), 0) as totalMileage"		'/사용마일리지
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by m.sitename " & groupsqlStr
		sqlStr = sqlStr & " 	order by yyyymm desc, m.sitename asc"
		sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<br>"
		'response.end
        ''-------------------------------------------------------------------------------------------------------------------
		if (FRectSearchField="sitename") and (FRectSearchText<>"") then
            FRectSitename=FRectSearchText
        end if

		sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','site','"&FRectDategbn&"'" & ", 0, 0, '" + CStr(FRectExcTPL) + "' "
        db3_rsget.CursorLocation = adUseClient
    	''db3_rsget.CursorType = adOpenStatic
    	''db3_rsget.LockType = adLockOptimistic
        ''db3_dbget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
'		db3_rsget.pagesize = FPageSize
'response.write sqlStr & "<br>"
'response.write "수정중..."
'response.end
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

'		if (FCurrPage * FPageSize < FTotalCount) then
'			FResultCount = FPageSize
'		else
'			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
'		end if
'
'		FTotalPage = (FTotalCount\FPageSize)
'		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
'
'		redim preserve FItemList(FResultCount)
'
'		FPageCount = FCurrPage - 1

        FResultCount = db3_rsget.Recordcount
        FTotalCount = FResultCount
        redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			'db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).ftotalMileage		= db3_rsget("totalMileage")
				FItemList(i).fsitename		= db3_rsget("sitename")
				FItemList(i).fyyyymm		= db3_rsget("yyyymm")
				FItemList(i).forgTotalPrice		= db3_rsget("orgTotalPrice")
				FItemList(i).fsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).ftotalsum		= db3_rsget("totalsum")
				FItemList(i).ftotalBeasongBonusCouponDiscount		= db3_rsget("totalBeasongBonusCouponDiscount")
				FItemList(i).ftotalBonusCouponDiscount		= db3_rsget("totalBonusCouponDiscount")
				FItemList(i).ftotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).fallatdiscountprice		= db3_rsget("allatdiscountprice")
				FItemList(i).ftotalMaechulPrice		= db3_rsget("totalMaechulPrice")
				FItemList(i).ftotalUpcheJungsanCash		= db3_rsget("totalUpcheJungsanCash")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end function

public FRectIsSum

	'/admin/maechul/maechul_month_brand_log.asp
	public function GetMaechul_month_brand_Log()
		dim sqlStr,i, addSqlStr, groupsqlStr, fieldsqlStr, indexmSqlStr, indexdSqlStr

		if FRectStartDate="" or FRectEndDate="" then exit function
		if FRectIsSum = "" then FRectIsSum = "N"
		if FRectDategbn="ActDate" then
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_actDate))"
		elseif FRectDategbn="chulgoDate" then
			indexdSqlStr = indexdSqlStr + " with (index(IX_tbl_order_detail_log_beasongdate))"
		else
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_ipkumdate))"
		end if
		if FRectDategbn="ActDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.actDate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.actDate,21) as yyyymm"
		elseif FRectDategbn="chulgoDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),d.beasongdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),d.beasongdate,21) as yyyymm"
		else
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.ipkumdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.ipkumdate,21) as yyyymm"
		end if

		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if FRectmwdiv_beasongdiv="M" or FRectmwdiv_beasongdiv="W" or FRectmwdiv_beasongdiv="U" or FRectmwdiv_beasongdiv="R" then
			addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + FRectmwdiv_beasongdiv + "'"
		elseif FRectmwdiv_beasongdiv="TT" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
		elseif FRectmwdiv_beasongdiv="UU" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
		end if
		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "'"
		end if
		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if
		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if


		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	d.makerid"
		sqlStr = sqlStr & " 	, d.vatinclude " & fieldsqlStr		'/과세구분
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by d.makerid, d.vatinclude " & groupsqlStr
		sqlStr = sqlStr & " 	having isnull(d.makerid,'')<>''"
		sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<br>"
		'response.end
'		db3_rsget.Open sqlStr,db3_dbget,1
'			FTotalCount = db3_rsget("cnt")
'		db3_rsget.Close
'
'		if FTotalCount<1 then exit function

		sqlStr = "select *"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " 	d.makerid"
		sqlStr = sqlStr & " 	, d.vatinclude " & fieldsqlStr		'/과세구분
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='M' then d.anbunAppliedPriceDetailSUM end),0) as totalMaechulPrice_M"		'/매입매출총액
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='W' then d.anbunAppliedPriceDetailSUM end),0) as totalMaechulPrice_W"		'/위탁매출총액
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='U' then d.anbunAppliedPriceDetailSUM end),0) as totalMaechulPrice_U"		'/업체매출총액
		sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.itemid=0 and left(d.itemoption, 1)='9' then d.anbunAppliedPriceDetailSUM end),0) as totalMaechulPrice_UU"		'/업배배송비
		sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.omwdiv='M' then d.anbunAppliedPriceDetailSUM-(d.upcheJungsanCash*d.itemno) end),0) as accountMaechulPrice_M"		'/매입회계매출
		sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.omwdiv='W' then d.anbunAppliedPriceDetailSUM-(d.upcheJungsanCash*d.itemno) end),0) as accountMaechulPrice_W"		'/위탁회계매출
		sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.omwdiv='U' then d.anbunAppliedPriceDetailSUM-(d.upcheJungsanCash*d.itemno) end),0) as accountMaechulPrice_U"		'/업체회계매출
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='M' then d.upcheJungsanCash*d.itemno end),0) as totalUpcheJungsanCash_M"		'/매입업체정산액
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='W' then d.upcheJungsanCash*d.itemno end),0) as totalUpcheJungsanCash_W"		'/위탁업체정산액
		sqlStr = sqlStr & " 	,isnull(sum(case when d.omwdiv='U' then d.upcheJungsanCash*d.itemno end),0) as totalUpcheJungsanCash_U"		'/업체업체정산액
	sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.itemid=0 and left(d.itemoption, 1)<>'9' then d.upcheJungsanCash end),0) as beasongUpcheJungsanCash_TT"		'/텐배배송비정산액
		sqlStr = sqlStr & " 	,isnull(sum(case"
		sqlStr = sqlStr & " 		when d.itemid=0 and left(d.itemoption, 1)='9' then d.upcheJungsanCash end),0) as beasongUpcheJungsanCash_UU"		'/업배배송비정산액
		sqlStr = sqlStr & " 	, IsNull(sum(d.mileage * d.itemno), 0) as totalMileage"		'/사용마일리지
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by d.makerid, d.vatinclude " & groupsqlStr
		sqlStr = sqlStr & " 	having isnull(d.makerid,'')<>''"
		sqlStr = sqlStr & " 	order by yyyymm desc, d.makerid asc"
		sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<br>"
		'response.end
        ''---------------------------------------------------------------------------------------------------------------------------------------
		if (FRectSearchField="sitename") and (FRectSearchText<>"") then
            FRectSitename=FRectSearchText
        end if

        sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_BrandSum_CNT_New] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','brand','"&FRectDategbn&"',"&CHKIIF(FRectExceptSite="on","1","0") & ", '" & CStr(FRectExcTPL)& "' , '"&CStr(FRectIsSum) & "' "

        db3_rsget.CursorLocation = adUseClient
    	'db3_rsget.CursorType = adOpenStatic
    	'db3_rsget.LockType = adLockOptimistic

		''response.write sqlStr &"<br>"
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

        if FTotalCount<1 then exit function

		sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_BrandSum_New] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','brand','"&FRectDategbn&"',"&Cstr(FPageSize * FCurrPage)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" & CStr(FRectExcTPL) & "' , '"&CStr(FRectIsSum)&"'"
   		db3_rsget.CursorLocation = adUseClient
    	'db3_rsget.CursorType = adOpenStatic
    	'db3_rsget.LockType = adLockOptimistic

		''response.write sqlStr &"<br>"
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)



		i=0

		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).ftotalMileage		= db3_rsget("totalMileage")
				FItemList(i).fyyyymm		= db3_rsget("yyyymm")
				FItemList(i).fmakerid		= db3_rsget("makerid")
				FItemList(i).fvatinclude		= db3_rsget("vatinclude")
				FItemList(i).ftotalMaechulPrice_M		= db3_rsget("totalMaechulPrice_M")
				FItemList(i).ftotalMaechulPrice_W		= db3_rsget("totalMaechulPrice_W")
				FItemList(i).ftotalMaechulPrice_U		= db3_rsget("totalMaechulPrice_U")
				FItemList(i).ftotalMaechulPrice_TT		= db3_rsget("totalMaechulPrice_TT")
				FItemList(i).ftotalMaechulPrice_UU		= db3_rsget("totalMaechulPrice_UU")
				FItemList(i).faccountMaechulPrice_M		= db3_rsget("accountMaechulPrice_M")
				FItemList(i).faccountMaechulPrice_W		= db3_rsget("accountMaechulPrice_W")
				FItemList(i).faccountMaechulPrice_U		= db3_rsget("accountMaechulPrice_U")
				FItemList(i).ftotalUpcheJungsanCash_M		= db3_rsget("totalUpcheJungsanCash_M")
				FItemList(i).ftotalUpcheJungsanCash_W		= db3_rsget("totalUpcheJungsanCash_W")
				FItemList(i).ftotalUpcheJungsanCash_U		= db3_rsget("totalUpcheJungsanCash_U")
				FItemList(i).fbeasongUpcheJungsanCash_TT		= db3_rsget("beasongUpcheJungsanCash_TT")
				FItemList(i).fbeasongUpcheJungsanCash_UU		= db3_rsget("beasongUpcheJungsanCash_UU")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end function

	'//admin/maechul/maechul_month_item_log.asp
    public function GetMaechul_month_item_Log_X()
	    dim i,sqlStr, addSqlStr, groupsqlStr, fieldsqlStr, indexmSqlStr, indexdSqlStr
        dim viewName

		if FRectStartDate="" or FRectEndDate="" then exit function

        if FRectStartDate <> "" then
			addSqlStr = addSqlStr + " and d.yyyymmdd>='" + CStr(FRectStartDate) + "'"
		end if
		if FRectEndDate <> "" then
			addSqlStr = addSqlStr + " and d.yyyymmdd<'" + CStr(FRectEndDate) + "'"
		end if

		groupsqlStr = groupsqlStr + " , convert(varchar(7),d.yyyymmdd)"
		fieldsqlStr = fieldsqlStr + " , convert(varchar(7),d.yyyymmdd) as yyyymm"

		if FRectDategbn="ActDate" then
		    viewName = "db_datamart.[dbo].[vw_orderLog_ipkumdate]"
		elseif FRectDategbn="chulgoDate" then
		    viewName = "db_datamart.[dbo].[vw_orderLog_ipkumdate]"
		else
		    viewName = "db_datamart.[dbo].[vw_orderLog_ipkumdate]"
	    end if


		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(d.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(d.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if FRectmwdiv_beasongdiv<>"" then
			addSqlStr = addSqlStr + " and d.mwdiv_beasongdiv='" + FRectmwdiv_beasongdiv + "'"
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
		    if (FRectSearchField="sitename") then
		        addSqlStr = addSqlStr + " and d.sitename= '" + CStr(FRectSearchText) + "'"
		    end if
		end if

		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and d.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if

		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " d.vatinclude " & fieldsqlStr
		sqlStr = sqlStr + " ,d.mwdiv_beasongdiv"
		sqlStr = sqlStr + " , IsNull(Sum(orgTotalPrice),0) as orgTotalPrice"		'/소비자가
		sqlStr = sqlStr + " , IsNull(Sum(subtotalpriceCouponNotApplied),0) as subtotalpriceCouponNotApplied"		'/판매가(할인가)
		sqlStr = sqlStr + " , IsNull(Sum(totalsum),0) as totalsum	"		'/상품쿠폰적용가
		sqlStr = sqlStr + " , IsNull(sum(totalBonusCouponDiscount),0) as totalBonusCouponDiscount"
		sqlStr = sqlStr + " , IsNull(sum(totalBeasongBonusCouponDiscount),0) as totalBeasongBonusCouponDiscount"
		sqlStr = sqlStr + " , IsNull(sum(totalPriceBonusCouponDiscount),0) as totalPriceBonusCouponDiscount"		'/정액쿠폰
		sqlStr = sqlStr + " , IsNull(sum(allatdiscountprice),0) as allatdiscountprice"		'/기타할인(올앳)
		sqlStr = sqlStr + " , IsNull(sum(totalMaechulPrice),0) as totalMaechulPrice"		'/매출총액
		sqlStr = sqlStr + " , IsNull((sum(totalBuycash)),0) as totalBuycash"
		sqlStr = sqlStr + " , IsNull(round(sum(totalBuycash),1),0) as totalBuycash"
		sqlStr = sqlStr + " , IsNull(Sum(totalBuycashVAT),0) as totalBuycashVAT"		'/공급가세액
		sqlStr = sqlStr + " , IsNull(sum(totalUpcheJungsanCash),0) as totalUpcheJungsanCash"		'/업체정산액
		sqlStr = sqlStr + " , IsNull(Sum(totalUpcheJungsanCashVAT),0) as totalUpcheJungsanCashVAT"		'/업체정산부가세
		sqlStr = sqlStr & " , IsNull(sum(totalMileage),0) as totalMileage"		'/사용마일리지
		sqlStr = sqlStr + " from "&viewName&" d with (nolock)  "
		sqlStr = sqlStr + " where 1=1 " & addSqlStr
		sqlStr = sqlStr + " group by d.vatinclude " & groupsqlStr
		sqlStr = sqlStr + " 	,d.mwdiv_beasongdiv"
		sqlStr = sqlStr + " order by yyyymm desc"

'rw sqlStr
'        sqlStr = "select top 100 convert(varchar(7),yyyymmdd) as yyyymm, *"
'        sqlStr = sqlStr + " from db_datamart.[dbo].[vw_orderLog_ipkumdate]"
'        sqlStr = sqlStr + " where yyyymmdd>='2013-12-01'"
'        sqlStr = sqlStr + " and yyyymmdd<'2013-12-11'"
'        sqlStr = sqlStr + " and targetGbn='OF'"
'        sqlStr = sqlStr + " group by convert(varchar(7),yyyymmdd)"

        if (FRectSearchField="sitename") and (FRectSearchText<>"") then
            FRectSitename=FRectSearchText
        end if

		sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"'"
        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.CursorType = adOpenStatic
    	db3_rsget.LockType = adLockOptimistic

'response.end

		db3_rsget.open sqlStr,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).ftotalMileage		= db3_rsget("totalMileage")
				FItemList(i).fmwdiv_beasongdiv		= db3_rsget("mwdiv_beasongdiv")
				FItemList(i).fyyyymm		= db3_rsget("yyyymm")
				FItemList(i).fvatinclude		= db3_rsget("vatinclude")
				FItemList(i).forgTotalPrice		= db3_rsget("orgTotalPrice")
				FItemList(i).fsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).ftotalsum		= db3_rsget("totalsum")
				FItemList(i).ftotalBonusCouponDiscount		= db3_rsget("totalBonusCouponDiscount")
				FItemList(i).ftotalBeasongBonusCouponDiscount		= db3_rsget("totalBeasongBonusCouponDiscount")
				FItemList(i).ftotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).fallatdiscountprice		= db3_rsget("allatdiscountprice")
				FItemList(i).ftotalMaechulPrice		= db3_rsget("totalMaechulPrice")
				FItemList(i).ftotalBuycash		= db3_rsget("totalBuycash")
				FItemList(i).ftotalBuycashVAT		= db3_rsget("totalBuycashVAT")
				FItemList(i).ftotalUpcheJungsanCash		= db3_rsget("totalUpcheJungsanCash")
				FItemList(i).ftotalUpcheJungsanCashVAT		= db3_rsget("totalUpcheJungsanCashVAT")

				db3_rsget.movenext
				i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	public function GetMaechul_month_item_Log_SUM()
		dim i, sqlStr

		'// -- exec db_analyze_data_raw.[dbo].[sp_TEN_OrderLog_ipkumdate_NoView_SUM] '2017-12'
		sqlStr = "select * from [db_analyze_data_raw].[dbo].[tbl_meachul_log_summary] with (nolock) "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + " and yyyymm >= '" & FRectStartDate & "' "
		sqlStr = sqlStr + " and yyyymm <= '" & FRectEndDate & "' "
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				sqlStr = sqlStr + " and targetGbn in ('ON','AC')"
			else
				sqlStr = sqlStr + " and targetGbn = '" + FRecttargetGbn + "'"
			end if
		end if
		sqlStr = sqlStr + " order by yyyymm desc, targetGbn desc, sitename asc, vatinclude asc, mwdiv_beasongdiv  asc, beadaldiv "
		''response.write sqlStr & "<br>"
		''response.End

		rsAnalget.CursorLocation = adUseClient
		rsAnalget.open sqlStr,dbAnalget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsAnalget.recordcount
		FresultCount = rsAnalget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsAnalget.Eof Then
			Do Until rsAnalget.Eof
				set FItemList(i) = new CMaechulLogItem

				if (FRectDategbn <> "actDateAndChulgoDate") then
					FItemList(i).fyyyymm		= rsAnalget("yyyymm")
				else
					FItemList(i).Fyyyymm		= rsAnalget("Actyyyymm")
					FItemList(i).Fyyyymm2		= rsAnalget("Dlvyyyymm")
				end If

				FItemList(i).Fsitename			= rsAnalget("sitename")
				FItemList(i).Fitemno			= rsAnalget("itemno")
				FItemList(i).Fbeadaldiv			= rsAnalget("beadaldiv")

				''FItemList(i).Fsitename			= rsAnalget("sitename")
				FItemList(i).ftotalMileage		= rsAnalget("totalMileage")
				FItemList(i).fmwdiv_beasongdiv	= rsAnalget("mwdiv_beasongdiv")
				FItemList(i).fvatinclude		= rsAnalget("vatinclude")
				FItemList(i).forgTotalPrice		= rsAnalget("orgTotalPrice")
				FItemList(i).fsubtotalpriceCouponNotApplied		= rsAnalget("subtotalpriceCouponNotApplied")
				FItemList(i).ftotalsum							= rsAnalget("totalsum")
				FItemList(i).ftotalBonusCouponDiscount			= rsAnalget("totalBonusCouponDiscount")
				FItemList(i).ftotalBeasongBonusCouponDiscount	= rsAnalget("totalBeasongBonusCouponDiscount")
				FItemList(i).ftotalPriceBonusCouponDiscount		= rsAnalget("totalPriceBonusCouponDiscount")
				FItemList(i).fallatdiscountprice				= rsAnalget("allatdiscountprice")
				FItemList(i).ftotalMaechulPrice					= rsAnalget("totalMaechulPrice")
				FItemList(i).ftotalBuycash						= rsAnalget("totalBuycash")
				FItemList(i).ftotalBuycashVAT					= rsAnalget("totalBuycashVAT")
				FItemList(i).ftotalUpcheJungsanCash				= rsAnalget("totalUpcheJungsanCash")
				FItemList(i).ftotalUpcheJungsanCashVAT			= rsAnalget("totalUpcheJungsanCashVAT")
				FItemList(i).FtargetGbn			= rsAnalget("targetGbn")

				if (FRectDategbn <> "actDateAndChulgoDate") then
					FItemList(i).FavgipgoPrice			= rsAnalget("avgipgoPrice")
					FItemList(i).FoverValueStockPrice	= rsAnalget("overValueStockPrice")
				end if

				rsAnalget.movenext
				i = i + 1
			Loop
		End If

		rsAnalget.close
	end function

    public function GetMaechul_month_item_Log()
	    dim i,sqlStr, addSqlStr, groupsqlStr, fieldsqlStr, indexmSqlStr, indexdSqlStr

		if FRectStartDate="" or FRectEndDate="" then exit Function

		If (FRectGrpBy = "") Then
			FRectGrpBy = "mw"
		End If

		if FRectDategbn="ActDate" then
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_actDate))"
		elseif FRectDategbn="chulgoDate" then
			indexdSqlStr = indexdSqlStr + " with (index(IX_tbl_order_detail_log_beasongdate))"
		else
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_ipkumdate))"
		end if
		if FRectDategbn="ActDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.actDate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.actDate,21) as yyyymm"
		elseif FRectDategbn="chulgoDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),d.beasongdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),d.beasongdate,21) as yyyymm"
		else
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.ipkumdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.ipkumdate,21) as yyyymm"
		end if
		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if FRectmwdiv_beasongdiv="M" or FRectmwdiv_beasongdiv="W" or FRectmwdiv_beasongdiv="U" or FRectmwdiv_beasongdiv="R" then
			addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + FRectmwdiv_beasongdiv + "'"
		elseif FRectmwdiv_beasongdiv="TT" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
		elseif FRectmwdiv_beasongdiv="UU" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
		end if
		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "'"
		end if
		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if

		if FRectExcTPL <> "" then
			addSqlStr = addSqlStr + " and m.sitename not in (select distinct id as sitename from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " d.vatinclude " & fieldsqlStr
		'sqlStr = sqlStr & " ,(case" & vbcrlf
		'sqlStr = sqlStr & " 	when d.itemid=0 and left(d.itemoption, 1)<>'9' then 'TT'" & vbcrlf
		'sqlStr = sqlStr & " 	when d.itemid=0 and left(d.itemoption, 1)='9' then 'UU'" & vbcrlf
		'sqlStr = sqlStr & " 	else d.omwdiv" & vbcrlf
		'sqlStr = sqlStr & " 	end) as mwdiv_beasongdiv" & vbcrlf
		sqlStr = sqlStr & " , d.omwdiv as mwdiv_beasongdiv" & vbcrlf
		sqlStr = sqlStr + " , IsNull(Sum(d.orgitemcost * d.itemno), 0) as orgTotalPrice"		'/소비자가
		sqlStr = sqlStr + " , IsNull(Sum(d.itemcostCouponNotApplied * d.itemno), 0) as subtotalpriceCouponNotApplied"		'/판매가(할인가)
		sqlStr = sqlStr + " , IsNull(Sum(d.itemcost * d.itemno), 0) as totalsum	"		'/상품쿠폰적용가
		sqlStr = sqlStr + " , IsNull(sum((d.itemcost - d.reducedPrice - IsNull(d.allAtDiscount, 0)) * d.itemno), 0) as totalBonusCouponDiscount"
		sqlStr = sqlStr + " , IsNull(sum((case when d.itemid = 0 then d.itemcost - d.reducedPrice else 0 end) * d.itemno), 0) as totalBeasongBonusCouponDiscount"
		sqlStr = sqlStr + " , IsNull(sum(d.anbunCouponPriceDetailSUM), 0) as totalPriceBonusCouponDiscount"		'/정액쿠폰
		sqlStr = sqlStr + " , IsNull(sum(IsNull(d.allAtDiscount, 0) * d.itemno), 0) as allatdiscountprice"		'/기타할인(올앳)
		sqlStr = sqlStr + " , IsNull(sum(d.anbunAppliedPriceDetailSUM), 0) as totalMaechulPrice"		'/매출총액
		sqlStr = sqlStr + " , IsNull(round((case "
		sqlStr = sqlStr + " 	when d.vatinclude='N' then sum(d.anbunAppliedPriceDetailSUM)"
		sqlStr = sqlStr + " 	when d.vatinclude='Y' then sum(d.anbunAppliedPriceDetailSUM)*10/11"
		sqlStr = sqlStr + " 	end),1), 0) as totalBuycash"
		'sqlStr = sqlStr + " , IsNull(Sum(d.buycash * d.itemno), 0) as totalBuycash"		'/공급가액
		sqlStr = sqlStr + " , IsNull(Sum(d.buycashVAT * d.itemno), 0) as totalBuycashVAT"		'/공급가세액
		sqlStr = sqlStr + " , IsNull(sum(d.upcheJungsanCash * d.itemno), 0) as totalUpcheJungsanCash"		'/업체정산액
		sqlStr = sqlStr + " , IsNull(Sum(d.upcheJungsanCashVAT * d.itemno), 0) as totalUpcheJungsanCashVAT"		'/업체정산부가세
		sqlStr = sqlStr & " , IsNull(sum(d.mileage * d.itemno), 0) as totalMileage"		'/사용마일리지
		sqlStr = sqlStr + " from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr + " join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + " 	on m.orderserial = d.orderserial"
		sqlStr = sqlStr + " 	and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr + " where 1=1 " & addSqlStr
		sqlStr = sqlStr + " group by d.vatinclude " & groupsqlStr
		'sqlStr = sqlStr & " 	,(case" & vbcrlf
		'sqlStr = sqlStr & " 		when d.itemid=0 and left(d.itemoption, 1)<>'9' then 'TT'" & vbcrlf
		'sqlStr = sqlStr & " 		when d.itemid=0 and left(d.itemoption, 1)='9' then 'UU'" & vbcrlf
		'sqlStr = sqlStr & " 		else d.omwdiv" & vbcrlf
		'sqlStr = sqlStr & " 		end)" & vbcrlf
		sqlStr = sqlStr & " , d.omwdiv" & vbcrlf
		sqlStr = sqlStr + " order by yyyymm desc"

		'response.write sqlStr & "<Br>"
		'response.end
        '' 신규 ------------------------------------------------------------------------------------
		if (FRectSearchField="sitename") and (FRectSearchText<>"") then
            FRectSitename=FRectSearchText
        end if

		if (FRectDategbn <> "actDateAndChulgoDate") then
			''sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "
			''''2015/12/31 수정
			sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate_NoView] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','" & FRectGrpBy & "','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "

            ''if (FRectDategbn="chulgoDate") then FRectDategbn="beasongdate"  ''2016/10/04
			'// 2016-02-23, skyer9
			If (FRectPurchaseType = "") Then
				FRectPurchaseType = 0
			End If

			sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate_NoView_TEST] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','" & FRectGrpBy & "','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "'," & FRectPurchaseType & ", '"+FRectNextMonthJungsanFixed+"'"
		else
			'sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_actdate_chulgodate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"', '"&CStr(FRectStartDate2)&"','"&CStr(FRectEndDate2)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "

			sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_actdate_chulgodate_NoView] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"', '"&CStr(FRectStartDate2)&"','"&CStr(FRectEndDate2)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "

		end if
'rw sqlStr
'response.write sqlStr & "<br>"
'response.write "수정중..."
'response.End

		If (FRectUseNewDB <> "") Then

		    if (FRectGrpBy="cate") then
		    ' sqlStr = Replace(sqlStr, "db_datamart.", "db_analyze_data_raw.")
		    ' response.write sqlStr & "<br>"
			' response.end
				if (FRectDategbn <> "actDateAndChulgoDate") then
					sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','" & FRectGrpBy & "','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "'," & FRectPurchaseType & " "
				else
					sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_actdate_chulgodate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"', '"&CStr(FRectStartDate2)&"','"&CStr(FRectEndDate2)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "
				end if

                response.Write sqlStr & "<Br>"
    			rsSTSget.CursorLocation = adUseClient
    			rsSTSget.open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly

    			FTotalCount = rsSTSget.recordcount
    			FresultCount = rsSTSget.recordcount

    			redim FItemList(FresultCount)
    			i = 0
    			If Not rsSTSget.Eof Then
    				Do Until rsSTSget.Eof
    					set FItemList(i) = new CMaechulLogItem

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).fyyyymm		= rsSTSget("yyyymm")
    					else
    						FItemList(i).Fyyyymm		= rsSTSget("Actyyyymm")
    						FItemList(i).Fyyyymm2		= rsSTSget("Dlvyyyymm")
    					end If

    					If (FRectGrpBy = "brand2") Then
    						FItemList(i).Fmakerid			= rsSTSget("makerid")
    						FItemList(i).Fpurchasetype		= rsSTSget("purchasetype")
							FItemList(i).fpurchasetypename		= rsSTSget("purchasetypename")
    					elseIf (FRectGrpBy = "cate") Then
    						FItemList(i).Fcatename			= rsSTSget("catename")
                            FItemList(i).Fsku				= rsSTSget("sku")
    					elseIf (FRectGrpBy = "mwch") Then
    						FItemList(i).Fsitename			= rsSTSget("sitename")
    						FItemList(i).Fbeadaldiv			= rsSTSget("beadaldiv")
							
    					elseIf (FRectGrpBy = "mwordr") Then
    						FItemList(i).Forderserial		= rsSTSget("orderserial")
    					Else
    						FItemList(i).Fsitename			= rsSTSget("sitename")
    					End If

    					''FItemList(i).Fsitename			= rsSTSget("sitename")
    					FItemList(i).Fitemno			= rsSTSget("itemno")
						FItemList(i).ftotalMileage		= rsSTSget("totalMileage")
    					FItemList(i).fmwdiv_beasongdiv	= rsSTSget("mwdiv_beasongdiv")
    					FItemList(i).fvatinclude		= rsSTSget("vatinclude")
    					FItemList(i).forgTotalPrice		= rsSTSget("orgTotalPrice")
    					FItemList(i).fsubtotalpriceCouponNotApplied		= rsSTSget("subtotalpriceCouponNotApplied")
    					FItemList(i).ftotalsum							= rsSTSget("totalsum")
    					FItemList(i).ftotalBonusCouponDiscount			= rsSTSget("totalBonusCouponDiscount")
    					FItemList(i).ftotalBeasongBonusCouponDiscount	= rsSTSget("totalBeasongBonusCouponDiscount")
    					FItemList(i).ftotalPriceBonusCouponDiscount		= rsSTSget("totalPriceBonusCouponDiscount")
    					FItemList(i).fallatdiscountprice				= rsSTSget("allatdiscountprice")
    					FItemList(i).ftotalMaechulPrice					= rsSTSget("totalMaechulPrice")
    					FItemList(i).ftotalBuycash						= rsSTSget("totalBuycash")
    					FItemList(i).ftotalBuycashVAT					= rsSTSget("totalBuycashVAT")
    					FItemList(i).ftotalUpcheJungsanCash				= rsSTSget("totalUpcheJungsanCash")
    					FItemList(i).ftotalUpcheJungsanCashVAT			= rsSTSget("totalUpcheJungsanCashVAT")
    					FItemList(i).FtargetGbn			= rsSTSget("targetGbn")

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).FavgipgoPrice			= rsSTSget("avgipgoPrice")
    						FItemList(i).FoverValueStockPrice	= rsSTSget("overValueStockPrice")
    					end if

    					rsSTSget.movenext
    					i = i + 1
    				Loop
    			End If

    			rsSTSget.close
		    elseif (FRectGrpBy="cateChn") then

				if (FRectDategbn <> "actDateAndChulgoDate") then
					sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','" & FRectGrpBy & "','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "'," & FRectPurchaseType & " "
				else
					sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_actdate_chulgodate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"', '"&CStr(FRectStartDate2)&"','"&CStr(FRectEndDate2)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "
				end if
    			rsSTSget.CursorLocation = adUseClient
    			rsSTSget.open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly

    			FTotalCount = rsSTSget.recordcount
    			FresultCount = rsSTSget.recordcount

    			redim FItemList(FresultCount)
    			i = 0
    			If Not rsSTSget.Eof Then
    				Do Until rsSTSget.Eof
    					set FItemList(i) = new CMaechulLogItem

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).fyyyymm		= rsSTSget("yyyymm")
    					else
    						FItemList(i).Fyyyymm		= rsSTSget("Actyyyymm")
    						FItemList(i).Fyyyymm2		= rsSTSget("Dlvyyyymm")
    					end If

						FItemList(i).Fcatename			= rsSTSget("catename")
						FItemList(i).FchannelName		= right(rsSTSget("chnName"),len(rsSTSget("chnName"))-2)
						FItemList(i).Fitemno			= rsSTSget("itemno")
                        FItemList(i).Fsku				= rsSTSget("sku")

    					''FItemList(i).Fsitename			= rsSTSget("sitename")
    					FItemList(i).ftotalMileage		= rsSTSget("totalMileage")
    					FItemList(i).fmwdiv_beasongdiv	= rsSTSget("mwdiv_beasongdiv")
    					FItemList(i).fvatinclude		= rsSTSget("vatinclude")
    					FItemList(i).forgTotalPrice		= rsSTSget("orgTotalPrice")
    					FItemList(i).fsubtotalpriceCouponNotApplied		= rsSTSget("subtotalpriceCouponNotApplied")
    					FItemList(i).ftotalsum							= rsSTSget("totalsum")
    					FItemList(i).ftotalBonusCouponDiscount			= rsSTSget("totalBonusCouponDiscount")
    					FItemList(i).ftotalBeasongBonusCouponDiscount	= rsSTSget("totalBeasongBonusCouponDiscount")
    					FItemList(i).ftotalPriceBonusCouponDiscount		= rsSTSget("totalPriceBonusCouponDiscount")
    					FItemList(i).fallatdiscountprice				= rsSTSget("allatdiscountprice")
    					FItemList(i).ftotalMaechulPrice					= rsSTSget("totalMaechulPrice")
    					FItemList(i).ftotalBuycash						= rsSTSget("totalBuycash")
    					FItemList(i).ftotalBuycashVAT					= rsSTSget("totalBuycashVAT")
    					FItemList(i).ftotalUpcheJungsanCash				= rsSTSget("totalUpcheJungsanCash")
    					FItemList(i).ftotalUpcheJungsanCashVAT			= rsSTSget("totalUpcheJungsanCashVAT")
    					FItemList(i).FtargetGbn			= rsSTSget("targetGbn")

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).FavgipgoPrice			= rsSTSget("avgipgoPrice")
    						FItemList(i).FoverValueStockPrice	= rsSTSget("overValueStockPrice")
    					end if

    					rsSTSget.movenext
    					i = i + 1
    				Loop
    			End If

    			rsSTSget.close
		    else
    		    ''statistics DB 사용
    		    if (FRectDategbn <> "actDateAndChulgoDate") then
    		        If (FRectPurchaseType = "") Then
        				FRectPurchaseType = 0
        			End If

        			sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','" & FRectGrpBy & "','"&FRectDategbn&"',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "'," & FRectPurchaseType & " ,'"+FRectNextMonthJungsanFixed+"'"
        		else
        			sqlStr = "exec db_statistics.[dbo].[sp_TEN_OrderLog_actdate_chulgodate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"', '"&CStr(FRectStartDate2)&"','"&CStr(FRectEndDate2)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','mw',"&CHKIIF(FRECTExceptDlv="on",1,0)&","&CHKIIF(FRectExceptSite="on","1","0") & ", '" + CStr(FRectExcTPL) + "' "

        		end if
    	'	rw sqlStr
    	response.write sqlStr & "<br>"
    	'response.end
    			'// 73번 디비 사용
    			''sqlStr = Replace(sqlStr, "db_datamart.", "db_analyze_data_raw.")

    			rsSTSget.CursorLocation = adUseClient
    			rsSTSget.open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly

    			FTotalCount = rsSTSget.recordcount
    			FresultCount = rsSTSget.recordcount

    			redim FItemList(FresultCount)
    			i = 0
    			If Not rsSTSget.Eof Then
    				Do Until rsSTSget.Eof
    					set FItemList(i) = new CMaechulLogItem

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).fyyyymm		= rsSTSget("yyyymm")
    					else
    						FItemList(i).Fyyyymm		= rsSTSget("Actyyyymm")
    						FItemList(i).Fyyyymm2		= rsSTSget("Dlvyyyymm")
    					end If

    					If (FRectGrpBy = "brand2") Then
    						FItemList(i).Fmakerid			= rsSTSget("makerid")
    						FItemList(i).Fpurchasetype		= rsSTSget("purchasetype")
							FItemList(i).fpurchasetypename		= rsSTSget("purchasetypename")
    					elseIf (FRectGrpBy = "cate") Then
    						FItemList(i).Fcatename			= rsSTSget("catename")
    					elseIf (FRectGrpBy = "mwch") Then
    						FItemList(i).Fsitename			= rsSTSget("sitename")
    						FItemList(i).Fbeadaldiv			= rsSTSget("beadaldiv")
    					elseIf (FRectGrpBy = "mwordr") Then
    						FItemList(i).Forderserial		= rsSTSget("orderserial")
    					Else
    						FItemList(i).Fsitename			= rsSTSget("sitename")
    					End If

    					''FItemList(i).Fsitename			= rsSTSget("sitename")
    					FItemList(i).Fitemno			= rsSTSget("itemno")
						FItemList(i).ftotalMileage		= rsSTSget("totalMileage")
    					FItemList(i).fmwdiv_beasongdiv	= rsSTSget("mwdiv_beasongdiv")
    					FItemList(i).fvatinclude		= rsSTSget("vatinclude")
    					FItemList(i).forgTotalPrice		= rsSTSget("orgTotalPrice")
    					FItemList(i).fsubtotalpriceCouponNotApplied		= rsSTSget("subtotalpriceCouponNotApplied")
    					FItemList(i).ftotalsum							= rsSTSget("totalsum")
    					FItemList(i).ftotalBonusCouponDiscount			= rsSTSget("totalBonusCouponDiscount")
    					FItemList(i).ftotalBeasongBonusCouponDiscount	= rsSTSget("totalBeasongBonusCouponDiscount")
    					FItemList(i).ftotalPriceBonusCouponDiscount		= rsSTSget("totalPriceBonusCouponDiscount")
    					FItemList(i).fallatdiscountprice				= rsSTSget("allatdiscountprice")
    					FItemList(i).ftotalMaechulPrice					= rsSTSget("totalMaechulPrice")
    					FItemList(i).ftotalBuycash						= rsSTSget("totalBuycash")
    					FItemList(i).ftotalBuycashVAT					= rsSTSget("totalBuycashVAT")
    					FItemList(i).ftotalUpcheJungsanCash				= rsSTSget("totalUpcheJungsanCash")
    					FItemList(i).ftotalUpcheJungsanCashVAT			= rsSTSget("totalUpcheJungsanCashVAT")
    					FItemList(i).FtargetGbn			= rsSTSget("targetGbn")

    					if (FRectDategbn <> "actDateAndChulgoDate") then
    						FItemList(i).FavgipgoPrice			= rsSTSget("avgipgoPrice")
    						FItemList(i).FoverValueStockPrice	= rsSTSget("overValueStockPrice")
    					end if

    					rsSTSget.movenext
    					i = i + 1
    				Loop
    			End If

    			rsSTSget.close
    		end if
            ''rw sqlStr & "<br>"
		Else
			''rw sqlStr & "<br>"
rw  sqlStr
''response.end
			db3_rsget.CursorLocation = adUseClient
			''db3_rsget.CursorType = adOpenStatic
			''db3_rsget.LockType = adLockOptimistic

			''db3_dbget.CommandTimeout = 60  ''2015/12/31 (기본 30초)  ''주석처리 2016/04/05 다른페이지에 영향 있음.. //adOpenForwardOnly, adLockReadOnly
			db3_rsget.open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

			FTotalCount = db3_rsget.recordcount
			FresultCount = db3_rsget.recordcount

			redim FItemList(FresultCount)
			i = 0
			If Not db3_rsget.Eof Then
				Do Until db3_rsget.Eof
					set FItemList(i) = new CMaechulLogItem

					if (FRectDategbn <> "actDateAndChulgoDate") then
						FItemList(i).fyyyymm		= db3_rsget("yyyymm")
					else
						FItemList(i).Fyyyymm		= db3_rsget("Actyyyymm")
						FItemList(i).Fyyyymm2		= db3_rsget("Dlvyyyymm")
					end If

					If (FRectGrpBy = "brand2") Then
						FItemList(i).Fmakerid			= db3_rsget("makerid")
						FItemList(i).Fpurchasetype		= db3_rsget("purchasetype")
						FItemList(i).fpurchasetypename		= db3_rsget("purchasetypename")
						FItemList(i).Fitemno			= db3_rsget("itemno")
					elseIf (FRectGrpBy = "cate") Then
						FItemList(i).Fcatename			= db3_rsget("catename")
						FItemList(i).Fitemno			= db3_rsget("itemno")
					Else
						FItemList(i).Fsitename			= db3_rsget("sitename")
					End If

					''FItemList(i).Fsitename			= db3_rsget("sitename")
					FItemList(i).ftotalMileage		= db3_rsget("totalMileage")
					FItemList(i).fmwdiv_beasongdiv	= db3_rsget("mwdiv_beasongdiv")
					FItemList(i).fvatinclude		= db3_rsget("vatinclude")
					FItemList(i).forgTotalPrice		= db3_rsget("orgTotalPrice")
					FItemList(i).fsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
					FItemList(i).ftotalsum							= db3_rsget("totalsum")
					FItemList(i).ftotalBonusCouponDiscount			= db3_rsget("totalBonusCouponDiscount")
					FItemList(i).ftotalBeasongBonusCouponDiscount	= db3_rsget("totalBeasongBonusCouponDiscount")
					FItemList(i).ftotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
					FItemList(i).fallatdiscountprice				= db3_rsget("allatdiscountprice")
					FItemList(i).ftotalMaechulPrice					= db3_rsget("totalMaechulPrice")
					FItemList(i).ftotalBuycash						= db3_rsget("totalBuycash")
					FItemList(i).ftotalBuycashVAT					= db3_rsget("totalBuycashVAT")
					FItemList(i).ftotalUpcheJungsanCash				= db3_rsget("totalUpcheJungsanCash")
					FItemList(i).ftotalUpcheJungsanCashVAT			= db3_rsget("totalUpcheJungsanCashVAT")
					FItemList(i).FtargetGbn			= db3_rsget("targetGbn")

					if (FRectDategbn <> "actDateAndChulgoDate") then
						FItemList(i).FavgipgoPrice			= db3_rsget("avgipgoPrice")
						FItemList(i).FoverValueStockPrice	= db3_rsget("overValueStockPrice")
					end if

					db3_rsget.movenext
					i = i + 1
				Loop
			End If

			db3_rsget.close
		End If

        ''response.write sqlStr & "<br>"

	end function

	'//admin/maechul/maechul_log_sum.asp
    public function GetMaechulLogSum()
        dim i,sqlStr, addSqlStr, osqlStr

		addSqlStr = ""

        if FRectOrgPayStartDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectOrgPayStartDate) + "'"
		end if

		if FRectOrgPayEndDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectOrgPayEndDate) + "'"
		end if

        if FRectActDateStartDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectActDateStartDate) + "'"
		end if

		if FRectActDateEndDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectActDateEndDate) + "'"
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if

		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if

		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if

		if FRectExcTPL <> "" then
			addSqlStr = addSqlStr + " and m.sitename not in (select distinct id as sitename from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
		end if

'		sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
'		sqlStr = sqlStr + " from "
'		sqlStr = sqlStr + " 	(select convert(varchar(10),m.ipkumdate,21) as ipkumdate"
'		sqlStr = sqlStr + " 	    from db_datamart.dbo.tbl_order_master_log m "
'		sqlStr = sqlStr + " where "
'		sqlStr = sqlStr + " 	1 = 1 "
'		sqlStr = sqlStr + addSqlStr
'		sqlStr = sqlStr + " 	group by convert(varchar(10),m.ipkumdate,21) ) as T"
'
'		'rw sqlstr
'    	db3_rsget.Open sqlStr,db3_dbget,1
'			FTotalCount = db3_rsget("cnt")
'			FTotalPage = db3_rsget("totPg")
'		db3_rsget.Close
'
'		'지정페이지가 전체 페이지보다 클 때 함수종료
'		if CLng(FCurrPage)>CLng(FTotalPage) then
'			FResultCount = 0
'			exit function
'		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		if (FRectDategbn="ActDate") then
			sqlStr = sqlStr + " convert(varchar(10),m.actDate,21) as ipkumdate"
		else
    		sqlStr = sqlStr + " convert(varchar(10),m.ipkumdate,21) as ipkumdate"
    	end if
		sqlStr = sqlStr + " , IsNull(Sum(m.orgTotalPrice), 0) as orgTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.subtotalpriceCouponNotApplied), 0) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalsum), 0) as totalsum "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPrice), 0) as totalReducedPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPriceBeasongPay), 0) as totalReducedPriceBeasongPay "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalItemCouponDiscount), 0) as totalItemCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBonusCouponDiscount), 0) as totalBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBeasongBonusCouponDiscount), 0) as totalBeasongBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.allatdiscountprice), 0) as allatdiscountprice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulPrice), 0) as totalMaechulPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulVatPrice), 0) as totalMaechulVatPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.mileTotalPrice), 0) as mileTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.giftTotalPrice), 0) as giftTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.depositTotalPrice), 0) as depositTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycash), 0) as totalBuycash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashVAT), 0) as totalBuycashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashCouponNotApplied), 0) as totalBuycashCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCash), 0) as totalUpcheJungsanCash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCashVAT), 0) as totalUpcheJungsanCashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMileage), 0) as totalMileage "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m with (nolock) "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + addSqlStr

		if (FRectDategbn="ActDate") then
			sqlStr = sqlStr + " group by convert(varchar(10),m.actDate,21) "
		else
    		sqlStr = sqlStr + " group by convert(varchar(10),m.ipkumdate,21) "
    	end if

        if (FRectTargetGbn<>"") and (FRectDategbn<>"ActDate") and (FRectChkOnlyDiff<>"") then
            osqlStr = sqlStr

            if (FRectTargetGbn="ON") then
                sqlStr = " select A.*,B.realTotalsum,B.realSpendmileage "'', B.realGainmileage
                sqlStr = sqlStr + " from ("
                sqlStr = sqlStr + osqlStr
                sqlStr = sqlStr + ") A"
    		    sqlStr = sqlStr + "		left join ("
                sqlStr = sqlStr + "		select ipkumdate as ipkumdate1,sum(isNULL(totalitemcostsum,0)+isNULL(totalDlvPay,0)-isNULL(spendScoupon,0)-isNULL(discountEtc,0)) as realTotalsum"
                sqlStr = sqlStr + "		,sum(isNULL(spendmileage,0)) as realSpendmileage"
                'sqlStr = sqlStr + "     ,sum(isNULL(totalmileage,0)) as realGainmileage"
                sqlStr = sqlStr + "		from db_datamart.dbo.tbl_mkt_daily_totalsale with (nolock) "
                sqlStr = sqlStr + "		where 1=1"
                sqlStr = sqlStr + "		and ipkumdate>='"&FRectOrgPayStartDate&"'"
                sqlStr = sqlStr + "		and ipkumdate<'"&FRectOrgPayEndDate&"'"
                sqlStr = sqlStr + "		group by ipkumdate"
    		    sqlStr = sqlStr + "		) B on A.ipkumdate=B.ipkumdate1"
            end if

            if (FRectTargetGbn="AC") then
                sqlStr = " select A.*,B.realTotalsum,B.realSpendmileage " '', B.realGainmileage
                sqlStr = sqlStr + " from ("
                sqlStr = sqlStr + osqlStr
                sqlStr = sqlStr + ") A"
    		    sqlStr = sqlStr + "		left join ("
                sqlStr = sqlStr + " select convert(varchar(10),ipkumdate,21) as ipkumdate1"
            	sqlStr = sqlStr + " ,sum(subtotalprice+isNULL(miletotalPrice,0)) as realTotalsum"
            	sqlStr = sqlStr + " ,sum(isNULL(miletotalPrice,0)) as realSpendmileage"
            	'sqlStr = sqlStr + " ,sum(isNULL(totalmileage,0)) as realGainmileage"
            	sqlStr = sqlStr + " from [ACADEMYDB].db_academy.dbo.tbl_academy_order_master with (nolock)"
            	sqlStr = sqlStr + " where ipkumdate>='"&FRectOrgPayStartDate&"'"
            	sqlStr = sqlStr + " and ipkumdate<'"&FRectOrgPayEndDate&"'"
            	sqlStr = sqlStr + " and ipkumdiv>3"
            	sqlStr = sqlStr + " and cancelyn='N'"
            	sqlStr = sqlStr + " group by convert(varchar(10),ipkumdate,21) "
            	sqlStr = sqlStr + "		) B on A.ipkumdate=B.ipkumdate1"
	        end if

            if (FRectTargetGbn="OF") then
                sqlStr = " select A.*,B.realTotalsum,B.realSpendmileage "'', B.realGainmileage
                sqlStr = sqlStr + " from ("
                sqlStr = sqlStr + osqlStr
                sqlStr = sqlStr + ") A"
    		    sqlStr = sqlStr + "		left join ("
                sqlStr = sqlStr + "		select convert(varchar(10),s.shopregdate,21) as ipkumdate1,sum(isNULL(s.realsum,0) + isNULL(s.spendmile,0)) as realTotalsum"
                sqlStr = sqlStr + "		,sum(isNULL(s.spendmile,0)) as realSpendmileage"
                sqlStr = sqlStr + "		from [db_shop].[dbo].tbl_shopjumun_master s with (nolock)"
				sqlStr = sqlStr + "		join [db_partner].[dbo].tbl_partner p with (nolock) "
				sqlStr = sqlStr + "		on p.id = s.shopid "
				sqlStr = sqlStr + "		join [db_user].[dbo].tbl_user_c c with (nolock) "
				sqlStr = sqlStr + "		on c.userid=p.id "
                sqlStr = sqlStr + "		where 1=1"
				sqlStr = sqlStr + "		and p.userdiv = '501' "
				sqlStr = sqlStr + "		and c.userdiv = '21' "
                sqlStr = sqlStr + "		and shopregdate >= '"&FRectOrgPayStartDate&"'"
                sqlStr = sqlStr + "		and shopregdate < '"&FRectOrgPayEndDate&"'"
				sqlStr = sqlStr + " 	and cancelyn='N'"
                sqlStr = sqlStr + "		group by convert(varchar(10),shopregdate,21)"
    		    sqlStr = sqlStr + "		) B on A.ipkumdate=B.ipkumdate1"
            end if

		end if

        sqlStr = sqlStr + " order by ipkumdate desc"

		'response.write sqlStr & "<Br>"
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/05

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulLogItem

'				FItemList(i).Forderserial		= db3_rsget("orderserial")
'				FItemList(i).Fsuborderserial	= db3_rsget("suborderserial")
'				FItemList(i).FactDivCode		= db3_rsget("actDivCode")
'				FItemList(i).FactDate			= db3_rsget("actDate")
'				FItemList(i).Fjumundiv			= db3_rsget("jumundiv")
'				FItemList(i).Faccountdiv		= db3_rsget("accountdiv")

				FItemList(i).ForgTotalPrice						= db3_rsget("orgTotalPrice")
				FItemList(i).FsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).Ftotalsum							= db3_rsget("totalsum")
				FItemList(i).FtotalReducedPrice					= db3_rsget("totalReducedPrice")
				FItemList(i).FtotalReducedPriceBeasongPay		= db3_rsget("totalReducedPriceBeasongPay")
				FItemList(i).FtotalItemCouponDiscount			= db3_rsget("totalItemCouponDiscount")
				FItemList(i).FtotalBonusCouponDiscount			= db3_rsget("totalBonusCouponDiscount") - db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).FtotalBeasongBonusCouponDiscount	= db3_rsget("totalBeasongBonusCouponDiscount")

				FItemList(i).FtotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).Fallatdiscountprice				= db3_rsget("allatdiscountprice")
				FItemList(i).FtotalMaechulPrice					= db3_rsget("totalMaechulPrice")
				FItemList(i).FtotalMaechulVatPrice				= db3_rsget("totalMaechulVatPrice")
				FItemList(i).FmileTotalPrice					= db3_rsget("mileTotalPrice")
				FItemList(i).FgiftTotalPrice					= db3_rsget("giftTotalPrice")
				FItemList(i).FdepositTotalPrice					= db3_rsget("depositTotalPrice")
				FItemList(i).FtotalBuycash						= db3_rsget("totalBuycash")
				FItemList(i).FtotalBuycashVAT					= db3_rsget("totalBuycashVAT")
				FItemList(i).FtotalBuycashCouponNotApplied		= db3_rsget("totalBuycashCouponNotApplied")
				FItemList(i).FtotalUpcheJungsanCash				= db3_rsget("totalUpcheJungsanCash")
				FItemList(i).FtotalUpcheJungsanCashVAT			= db3_rsget("totalUpcheJungsanCashVAT")

				FItemList(i).FtotalMileage	= db3_rsget("totalMileage")
				FItemList(i).Fipkumdate		= db3_rsget("ipkumdate")
'				FItemList(i).Fsitename		= db3_rsget("sitename")
'				FItemList(i).Frdsite		= db3_rsget("rdsite")
'				FItemList(i).Fregdate		= db3_rsget("regdate")

                if (FRectDategbn<>"ActDate") and (FRectChkOnlyDiff<>"") and (FRectTargetGbn <> "") then
				    FItemList(i).FrealTotalsum	= db3_rsget("realTotalsum")
				    FItemList(i).FrealSpendmileage	= db3_rsget("realSpendmileage")
				    'FItemList(i).FrealGainmileage= db3_rsget("realGainmileage")
                end if

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	'//admin/maechul/maechul_log.asp
	public function GetMaechulLog()
	    dim i,sqlStr, addSqlStr

		addSqlStr = ""

		if FRectOrgPayStartDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectOrgPayStartDate) + "'"
		end if
		if FRectOrgPayEndDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectOrgPayEndDate) + "'"
		end if

		if FRectActDateStartDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectActDateStartDate) + "'"
		end if
		if FRectActDateEndDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectActDateEndDate) + "'"
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if

		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if

		if FRectExcTPL <> "" then
			addSqlStr = addSqlStr + " and m.sitename not in (select distinct id as sitename from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
		end if

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
			sqlStr = sqlStr + " from ( "
			sqlStr = sqlStr + " 	select "
			''sqlStr = sqlStr + " 		distinct m.orderserial, IsNull(Sum(m.totalMaechulPrice - (m.mileTotalPrice + m.giftTotalPrice + m.depositTotalPrice)), 0) as realpaysum "
			sqlStr = sqlStr + " 		distinct m.orderserial, IsNull(Sum(m.totalMaechulPrice), 0) as realpaysum "
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 		db_datamart.dbo.tbl_order_master_log m with (nolock) "
			sqlStr = sqlStr + " 	where "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + addSqlStr

			sqlStr = sqlStr + " 	group by orderserial, LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19), m.jumundiv, m.accountdiv, m.sitename, m.rdsite "
			sqlStr = sqlStr + " ) T "

			if (FRectChkOnlyDiff = "Y") then
				sqlStr = sqlStr + " left join db_order.dbo.tbl_order_master m with (nolock) "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
				sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
				sqlStr = sqlStr + " left join db_log.dbo.tbl_old_order_master_2003 om with (nolock) "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.orderserial = om.orderserial "
				sqlStr = sqlStr + " 	and om.cancelyn = 'N' "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				''sqlStr = sqlStr + " 	and (IsNull(m.subtotalprice, 0) - IsNull(m.sumPaymentEtc, 0) + IsNull(om.subtotalprice, 0) - IsNull(om.sumPaymentEtc, 0)) <>  T.realpaysum "
				sqlStr = sqlStr + " 	and (IsNull(m.subtotalprice, 0) + IsNull(m.miletotalprice, 0) + IsNull(om.subtotalprice, 0) + IsNull(om.miletotalprice, 0)) <>  T.realpaysum "
			end if

		else
			sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m with (nolock) "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + addSqlStr
		end if

		'rw sqlstr
		db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		sqlStr = "select "
		sqlStr = sqlStr + " IsNull(Sum(m.orgTotalPrice), 0) as orgTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.subtotalpriceCouponNotApplied), 0) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalsum), 0) as totalsum "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPrice), 0) as totalReducedPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPriceBeasongPay), 0) as totalReducedPriceBeasongPay "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalItemCouponDiscount), 0) as totalItemCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBonusCouponDiscount), 0) as totalBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBeasongBonusCouponDiscount), 0) as totalBeasongBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.allatdiscountprice), 0) as allatdiscountprice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulPrice), 0) as totalMaechulPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulVatPrice), 0) as totalMaechulVatPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.mileTotalPrice), 0) as mileTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.giftTotalPrice), 0) as giftTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.depositTotalPrice), 0) as depositTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycash), 0) as totalBuycash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashVAT), 0) as totalBuycashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashCouponNotApplied), 0) as totalBuycashCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCash), 0) as totalUpcheJungsanCash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCashVAT), 0) as totalUpcheJungsanCashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMileage), 0) as totalMileage "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m with (nolock) "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + addSqlStr

		''response.end
		db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
			Set FOneItem = new CMaechulLogSUMItem

			FOneItem.ForgTotalPrice = db3_rsget("orgTotalPrice")
			FOneItem.FsubtotalpriceCouponNotApplied = db3_rsget("subtotalpriceCouponNotApplied")
			FOneItem.Ftotalsum = db3_rsget("totalsum")
			FOneItem.FtotalReducedPrice = db3_rsget("totalReducedPrice")
			''FOneItem.FanbunPriceDetailSUM = db3_rsget("anbunPriceDetailSUM")
			FOneItem.FtotalMaechulPrice = db3_rsget("totalMaechulPrice")
			FOneItem.FtotalBuycash = db3_rsget("totalBuycash")

			FOneItem.FmileTotalPrice = db3_rsget("mileTotalPrice")
			FOneItem.FgiftTotalPrice = db3_rsget("giftTotalPrice")
			FOneItem.FdepositTotalPrice = db3_rsget("depositTotalPrice")

			FOneItem.FtotalUpcheJungsanCash = db3_rsget("totalUpcheJungsanCash")
			FOneItem.FtotalMileage = db3_rsget("totalMileage")
			FOneItem.FtotalBonusCouponDiscount = db3_rsget("totalBonusCouponDiscount")
			FOneItem.FtotalBeasongBonusCouponDiscount = db3_rsget("totalBeasongBonusCouponDiscount")
			FOneItem.FtotalPriceBonusCouponDiscount = db3_rsget("totalPriceBonusCouponDiscount")
			FOneItem.Fallatdiscountprice = db3_rsget("allatdiscountprice")

		db3_rsget.Close

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.orderserial "
			sqlStr = sqlStr + " , m.orderserial "
			sqlStr = sqlStr + " , '' as suborderserial "
			sqlStr = sqlStr + " , '' as actDivCode "
			sqlStr = sqlStr + " , LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19) as actDate "
			sqlStr = sqlStr + " , m.jumundiv "
			sqlStr = sqlStr + " , m.accountdiv "

			sqlStr = sqlStr + " , IsNull(Sum(m.orgTotalPrice), 0) as orgTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.subtotalpriceCouponNotApplied), 0) as subtotalpriceCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalsum), 0) as totalsum "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPrice), 0) as totalReducedPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPriceBeasongPay), 0) as totalReducedPriceBeasongPay "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalItemCouponDiscount), 0) as totalItemCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBonusCouponDiscount), 0) as totalBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBeasongBonusCouponDiscount), 0) as totalBeasongBonusCouponDiscount "

			sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.allatdiscountprice), 0) as allatdiscountprice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulPrice), 0) as totalMaechulPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulVatPrice), 0) as totalMaechulVatPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.mileTotalPrice), 0) as mileTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.giftTotalPrice), 0) as giftTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.depositTotalPrice), 0) as depositTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycash), 0) as totalBuycash "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashVAT), 0) as totalBuycashVAT "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashCouponNotApplied), 0) as totalBuycashCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCash), 0) as totalUpcheJungsanCash "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCashVAT), 0) as totalUpcheJungsanCashVAT "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMileage), 0) as totalMileage "

			sqlStr = sqlStr + " , LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19) as ipkumdate "
			sqlStr = sqlStr + " , m.sitename "
			sqlStr = sqlStr + " , m.rdsite "
			sqlStr = sqlStr + " , '' as regdate "
			''sqlStr = sqlStr + " , ((IsNull((select distinct (subtotalprice - sumPaymentEtc) from db_order.dbo.tbl_order_master where orderserial = m.orderserial and cancelyn <> 'Y'), 0) + IsNull((select distinct (subtotalprice - sumPaymentEtc) from db_log.dbo.tbl_old_order_master_2003 where orderserial = m.orderserial and cancelyn <> 'Y'), 0)) - IsNull((Sum(m.totalMaechulPrice) - (Sum(m.mileTotalPrice) + Sum(m.giftTotalPrice) + Sum(m.depositTotalPrice))), 0)) as realTotalsum "
			sqlStr = sqlStr + " , ((IsNull((select distinct (subtotalprice + IsNull(miletotalprice, 0)) from db_order.dbo.tbl_order_master where orderserial = m.orderserial and cancelyn = 'N'), 0) + IsNull((select distinct (subtotalprice + IsNull(miletotalprice, 0)) from db_log.dbo.tbl_old_order_master_2003 where orderserial = m.orderserial and cancelyn = 'N'), 0)) - IsNull(Sum(m.totalMaechulPrice), 0)) as realTotalsum "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m with (nolock) "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + addSqlStr

			sqlStr = sqlStr + "group by "
			sqlStr = sqlStr + "		orderserial, LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19), m.jumundiv, m.accountdiv, m.sitename, m.rdsite "

			if (FRectChkOnlyDiff = "Y") then
				sqlStr = sqlStr + " having "
				''sqlStr = sqlStr + "	(IsNull((select distinct (subtotalprice - sumPaymentEtc) from db_order.dbo.tbl_order_master where orderserial = m.orderserial and cancelyn <> 'Y'), 0) + IsNull((select distinct (subtotalprice - sumPaymentEtc) from db_log.dbo.tbl_old_order_master_2003 where orderserial = m.orderserial and cancelyn <> 'Y'), 0)) <> IsNull((Sum(m.totalMaechulPrice) - (Sum(m.mileTotalPrice) + Sum(m.giftTotalPrice) + Sum(m.depositTotalPrice))), 0) "
				''sqlStr = sqlStr + "	(IsNull((select distinct (subtotalprice + IsNull(miletotalprice, 0)) from db_order.dbo.tbl_order_master where orderserial = m.orderserial and cancelyn = 'N'), 0) + IsNull((select distinct (subtotalprice + IsNull(miletotalprice, 0)) from db_log.dbo.tbl_old_order_master_2003 where orderserial = m.orderserial and cancelyn = 'N'), 0)) <> IsNull(Sum(m.totalMaechulPrice), 0) "

				// 당일 취소내역 있는 경우 합쳐준다.
				sqlStr = sqlStr + "	IsNull((select "
				sqlStr = sqlStr + "		((case when m1.cancelyn = 'N' then m1.subtotalprice + IsNull(m1.miletotalprice, 0) else 0 end) "
				sqlStr = sqlStr + "		+ "
				sqlStr = sqlStr + "		IsNull(( "
				sqlStr = sqlStr + "			select sum(r1.canceltotal - r1.refundmileagesum) "
				sqlStr = sqlStr + "			from "
				sqlStr = sqlStr + "				db_cs.dbo.tbl_new_as_list a1 "
				sqlStr = sqlStr + "				join db_cs.dbo.tbl_as_refund_info r1 "
				sqlStr = sqlStr + "				on "
				sqlStr = sqlStr + "					a1.id = r1.asid "
				sqlStr = sqlStr + "			where "
				sqlStr = sqlStr + "				1 = 1 "
				sqlStr = sqlStr + "				and a1.orderserial = m1.orderserial "
				sqlStr = sqlStr + "				and a1.divcd = 'A008' "
				sqlStr = sqlStr + "				and a1.currstate = 'B007' "
				sqlStr = sqlStr + "				and a1.deleteyn = 'N' "
				sqlStr = sqlStr + "				and a1.finishdate >= convert(varchar(10), getdate(), 121) "
				sqlStr = sqlStr + "		), 0)) "
				sqlStr = sqlStr + "	from "
				sqlStr = sqlStr + "		db_order.dbo.tbl_order_master m1 with (nolock) "
				sqlStr = sqlStr + "	where "
				sqlStr = sqlStr + "		1 = 1 "
				sqlStr = sqlStr + "		and m1.orderserial = m.orderserial), 0) <> IsNull(Sum(m.totalMaechulPrice), 0) "
			end if

    		sqlStr = sqlStr + " order by LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19) desc"
		else
			sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.ipkumdate, 120), 120), 19) as ipkumdate, LEFT(CONVERT(VARCHAR, CONVERT(datetime, m.actDate, 120), 120), 19) as actDate, 0 as realTotalsum "
			sqlStr = sqlStr + " , IsNull(m.orgTotalPrice, 0) as orgTotalPrice, IsNull(m.subtotalpriceCouponNotApplied, 0) as subtotalpriceCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(m.totalItemCouponDiscount, 0) as totalItemCouponDiscount, IsNull(m.totalBuycashCouponNotApplied, 0) as totalBuycashCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(totalPriceBonusCouponDiscount, 0) as totalPriceBonusCouponDiscount, IsNull(allatdiscountprice, 0) as allatdiscountprice, IsNull(totalMaechulPrice, 0) as totalMaechulPrice"
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m with (nolock) "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + addSqlStr

    		sqlStr = sqlStr + " order by m.actDate desc"
		end if

		''response.write sqlStr & "<Br>"
		''response.end
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).Forderserial		= db3_rsget("orderserial")
				FItemList(i).Fsuborderserial	= db3_rsget("suborderserial")
				FItemList(i).FactDivCode		= db3_rsget("actDivCode")
				FItemList(i).FactDate			= db3_rsget("actDate")
				FItemList(i).Fjumundiv			= db3_rsget("jumundiv")
				FItemList(i).Faccountdiv		= db3_rsget("accountdiv")

				FItemList(i).ForgTotalPrice						= db3_rsget("orgTotalPrice")
				FItemList(i).FsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).Ftotalsum							= db3_rsget("totalsum")
				FItemList(i).FtotalReducedPrice					= db3_rsget("totalReducedPrice")
				FItemList(i).FtotalReducedPriceBeasongPay		= db3_rsget("totalReducedPriceBeasongPay")
				FItemList(i).FtotalItemCouponDiscount			= db3_rsget("totalItemCouponDiscount")
				FItemList(i).FtotalBonusCouponDiscount			= db3_rsget("totalBonusCouponDiscount") - db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).FtotalBeasongBonusCouponDiscount	= db3_rsget("totalBeasongBonusCouponDiscount")

				FItemList(i).FtotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).Fallatdiscountprice				= db3_rsget("allatdiscountprice")
				FItemList(i).FtotalMaechulPrice					= db3_rsget("totalMaechulPrice")
				FItemList(i).FtotalMaechulVatPrice				= db3_rsget("totalMaechulVatPrice")
				FItemList(i).FmileTotalPrice					= db3_rsget("mileTotalPrice")
				FItemList(i).FgiftTotalPrice					= db3_rsget("giftTotalPrice")
				FItemList(i).FdepositTotalPrice					= db3_rsget("depositTotalPrice")
				FItemList(i).FtotalBuycash						= db3_rsget("totalBuycash")
				FItemList(i).FtotalBuycashVAT					= db3_rsget("totalBuycashVAT")
				FItemList(i).FtotalBuycashCouponNotApplied		= db3_rsget("totalBuycashCouponNotApplied")
				FItemList(i).FtotalUpcheJungsanCash				= db3_rsget("totalUpcheJungsanCash")
				FItemList(i).FtotalUpcheJungsanCashVAT			= db3_rsget("totalUpcheJungsanCashVAT")

				FItemList(i).FtotalMileage	= db3_rsget("totalMileage")
				FItemList(i).Fipkumdate		= db3_rsget("ipkumdate")
				FItemList(i).Fsitename		= db3_rsget("sitename")
				FItemList(i).Frdsite		= db3_rsget("rdsite")
				FItemList(i).Fregdate		= db3_rsget("regdate")

				FItemList(i).FrealTotalsum	= db3_rsget("realTotalsum")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	'//admin/maechul/maechul_log.asp
	public function GetMaechulLogByMonth()
	    dim i,sqlStr, addSqlStr, tmpSqlStr, addLoopSqlStr, sqlUnionStr
		Dim dateField
		Dim yyyy, mm, dd, tmpStartDate, tmpEndDate, tmpDate

		addSqlStr = ""

		Select Case FRectDategbn
			Case "PayDate"
				dateField = "m.ipkumdate"
			Case Else
				dateField = "m.actDate"
		End Select

		addSqlStr = addSqlStr + " and  " & dateField & ">='" + CStr(FRectStartDate) + "'"
		addSqlStr = addSqlStr + " and " & dateField & "<'" + CStr(FRectEndDate) + "'"

		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRecttargetGbn <> "" then
			addSqlStr = addSqlStr + " and m.targetGbn = '" + FRectTargetGbn + "'"
		end if
		if FRectvatinclude<>"" then
			addSqlStr = addSqlStr + " and exists(select orderserial from db_datamart.dbo.tbl_order_detail_log as d1 where d1.orderserial=m.orderserial and d1.vatinclude='" + FRectvatinclude + "')"
		end if
		if FRectmwdiv_beasongdiv<>"" then
			addSqlStr = addSqlStr + " and exists(select orderserial from db_datamart.dbo.tbl_order_detail_log as d2 where d2.orderserial=m.orderserial and d2.omwdiv='" + FRectmwdiv_beasongdiv + "')"
		end if
		if FRectPurchasetype<>"" then
			if FRectPurchasetype="101" then
				addSqlStr = addSqlStr + " and exists(select orderserial from db_datamart.dbo.tbl_order_detail_log as d3 join db_partner.dbo.tbl_partner as p1 on d3.makerid=p1.id where d3.orderserial=m.orderserial and p1.purchasetype<>1)"
			elseif FRectPurchasetype="102" then
				addSqlStr = addSqlStr + " and exists(select orderserial from db_datamart.dbo.tbl_order_detail_log as d3 join db_partner.dbo.tbl_partner as p1 on d3.makerid=p1.id where d3.orderserial=m.orderserial and p1.purchasetype in (3,5,6))"
			else
				addSqlStr = addSqlStr + " and exists(select orderserial from db_datamart.dbo.tbl_order_detail_log as d3 join db_partner.dbo.tbl_partner as p1 on d3.makerid=p1.id where d3.orderserial=m.orderserial and p1.purchasetype='" + FRectPurchasetype + "')"
			end if
		end if

		''if FRectExcTPL <> "" then
		''	addSqlStr = addSqlStr + " and m.sitename not in (select distinct id as sitename from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
		''end If

		sqlStr = ""
		sqlUnionStr = ""

		yyyy = Left(FRectStartDate,4)
		mm = Right(Left(FRectStartDate,7),2)
		dd = Right(FRectStartDate,2)
		tmpStartDate = DateSerial(yyyy,mm,dd)

		yyyy = Left(FRectEndDate,4)
		mm = Right(Left(FRectEndDate,7),2)
		dd = Right(FRectEndDate,2)
		tmpEndDate = DateSerial(yyyy,mm,dd-1)

		''response.Write Left(tmpStartDate,10) + "<br>"
		''response.Write Left(tmpEndDate,10) + "<br>"

		Do until (Left(tmpStartDate,10) > Left(tmpEndDate,10))
			addLoopSqlStr = ""

			addLoopSqlStr = addLoopSqlStr + " and  " & dateField & ">='" + Left(tmpStartDate,10) + "'"
			If (Left(tmpStartDate,7) = Left(tmpEndDate,7)) Then
				tmpDate = DateAdd("d",1,tmpEndDate)
			Else
				tmpDate = DateSerial(Year(tmpStartDate), Month(tmpStartDate)+1,1)
			End If
			addLoopSqlStr = addLoopSqlStr + " and " & dateField & "<'" + Left(tmpDate,10) + "'"

			sqlStr = "select "
			sqlStr = sqlStr + " '" & Left(tmpStartDate,7) & "' as yyyymm "
			sqlStr = sqlStr + " , m.sitename "
			sqlStr = sqlStr + " , m.beadaldiv "
			sqlStr = sqlStr + " , sum(case when m.actDivCode = 'A' then 1 else 0 end) as orgOrderCnt "
			sqlStr = sqlStr + " , sum(case when m.actDivCode = 'C' then 1 when m.actDivCode = 'CC' then -1 else 0 end) as cancelOrderCnt "
			sqlStr = sqlStr + " , sum(case when m.actDivCode = 'M' then 1 when m.actDivCode = 'MM' then -1 else 0 end) as returnOrderCnt "
			sqlStr = sqlStr + " , IsNull(Sum(m.orgTotalPrice), 0) as orgTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.subtotalpriceCouponNotApplied), 0) as subtotalpriceCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalsum), 0) as totalsum "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPrice), 0) as totalReducedPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPriceBeasongPay), 0) as totalReducedPriceBeasongPay "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalItemCouponDiscount), 0) as totalItemCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBonusCouponDiscount), 0) as totalBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBeasongBonusCouponDiscount), 0) as totalBeasongBonusCouponDiscount "
			sqlStr = sqlStr + " , IsNull(Sum(m.allatdiscountprice), 0) as allatdiscountprice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulPrice), 0) as totalMaechulPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulVatPrice), 0) as totalMaechulVatPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.mileTotalPrice), 0) as mileTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.giftTotalPrice), 0) as giftTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.depositTotalPrice), 0) as depositTotalPrice "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycash), 0) as totalBuycash "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashVAT), 0) as totalBuycashVAT "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashCouponNotApplied), 0) as totalBuycashCouponNotApplied "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCash), 0) as totalUpcheJungsanCash "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCashVAT), 0) as totalUpcheJungsanCashVAT "
			sqlStr = sqlStr + " , IsNull(Sum(m.totalMileage), 0) as totalMileage "
            if (FRectShowLevel = "Y") then
                ''sqlStr = sqlStr + " , IsNull(r.userlevel, '-1') as userlevel "
                sqlStr = sqlStr + " , (case "
                sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.beadaldiv in (50, 51) then m.beadaldiv "
                sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.userid = '' then 99 "
                sqlStr = sqlStr + "     else IsNull(IsNull(r.userlevel, r2.userlevel), '-1') end) as userlevel "
            end if
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m with (nolock) "
			if (FRectShowLevel = "Y") then
				sqlStr = sqlStr + " left join [db_replica].[dbo].[tbl_order_master] r "
				sqlStr = sqlStr + " on 1=1 "
				sqlStr = sqlStr + " and m.orderserial = r.orderserial "
				sqlStr = sqlStr + " and m.targetGbn = 'ON' "
				sqlStr = sqlStr + " left join [db_log].[dbo].[tbl_old_order_master_2003] r2 "
				sqlStr = sqlStr + " on 1=1 "
				sqlStr = sqlStr + " and m.orderserial = r2.orderserial "
				sqlStr = sqlStr + " and m.targetGbn = 'ON' "
			end if
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + "		1 = 1 "
			sqlStr = sqlStr + addSqlStr
			sqlStr = sqlStr + addLoopSqlStr

			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	m.sitename "
			sqlStr = sqlStr + " 	, m.beadaldiv "
			if (FRectShowLevel = "Y") then
				''sqlStr = sqlStr + " , IsNull(r.userlevel, '-1') "
                sqlStr = sqlStr + " , (case "
                sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.beadaldiv in (50, 51) then m.beadaldiv "
                sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.userid = '' then 99 "
                sqlStr = sqlStr + "     else IsNull(IsNull(r.userlevel, r2.userlevel), '-1') end) "
			end if

			If (sqlUnionStr <> "") Then
				sqlUnionStr = sqlUnionStr + " union all "
			End If

			sqlUnionStr = sqlUnionStr + sqlStr

			tmpStartDate = tmpDate
		Loop

		''response.Write sqlUnionStr
		''response.End

		sqlStr = "select "
		sqlStr = sqlStr + " convert(varchar(7), " & dateField & ", 121) as yyyymm "
		sqlStr = sqlStr + " , m.sitename "
		sqlStr = sqlStr + " , m.beadaldiv "
		sqlStr = sqlStr + " , sum(case when m.actDivCode = 'A' then 1 else 0 end) as orgOrderCnt "
		sqlStr = sqlStr + " , sum(case when m.actDivCode = 'C' then 1 when m.actDivCode = 'CC' then -1 else 0 end) as cancelOrderCnt "
		sqlStr = sqlStr + " , sum(case when m.actDivCode = 'M' then 1 when m.actDivCode = 'MM' then -1 else 0 end) as returnOrderCnt "
		sqlStr = sqlStr + " , IsNull(Sum(m.orgTotalPrice), 0) as orgTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.subtotalpriceCouponNotApplied), 0) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalsum), 0) as totalsum "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPrice), 0) as totalReducedPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalReducedPriceBeasongPay), 0) as totalReducedPriceBeasongPay "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalItemCouponDiscount), 0) as totalItemCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBonusCouponDiscount), 0) as totalBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalPriceBonusCouponDiscount), 0) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBeasongBonusCouponDiscount), 0) as totalBeasongBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(Sum(m.allatdiscountprice), 0) as allatdiscountprice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulPrice), 0) as totalMaechulPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMaechulVatPrice), 0) as totalMaechulVatPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.mileTotalPrice), 0) as mileTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.giftTotalPrice), 0) as giftTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.depositTotalPrice), 0) as depositTotalPrice "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycash), 0) as totalBuycash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashVAT), 0) as totalBuycashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalBuycashCouponNotApplied), 0) as totalBuycashCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCash), 0) as totalUpcheJungsanCash "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalUpcheJungsanCashVAT), 0) as totalUpcheJungsanCashVAT "
		sqlStr = sqlStr + " , IsNull(Sum(m.totalMileage), 0) as totalMileage "
        if (FRectShowLevel = "Y") then
            ''sqlStr = sqlStr + " , IsNull(r.userlevel, '-1') as userlevel "
            sqlStr = sqlStr + " , (case "
            sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.beadaldiv in (50, 51) then m.beadaldiv "
            sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.userid = '' then 99 "
            sqlStr = sqlStr + "     else IsNull(IsNull(r.userlevel, r2.userlevel), '-1') end) as userlevel "
        end if
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m with (nolock) "
		if (FRectShowLevel = "Y") then
			sqlStr = sqlStr + " left join [db_replica].[dbo].[tbl_order_master] r1 "
			sqlStr = sqlStr + " on 1=1 "
			sqlStr = sqlStr + " and m.orderserial = r.orderserial "
			sqlStr = sqlStr + " and m.targetGbn = 'ON' "
			sqlStr = sqlStr + " left join [db_log].[dbo].[tbl_old_order_master_2003] r2 "
			sqlStr = sqlStr + " on 1=1 "
			sqlStr = sqlStr + " and m.orderserial = r2.orderserial "
			sqlStr = sqlStr + " and m.targetGbn = 'ON' "
		end if
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + addSqlStr

		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	convert(varchar(7), " & dateField & ", 121) "
		sqlStr = sqlStr + " 	, m.sitename "
		sqlStr = sqlStr + " 	, m.beadaldiv "
		if (FRectShowLevel = "Y") then
			''sqlStr = sqlStr + " , IsNull(r.userlevel, '-1') "
            sqlStr = sqlStr + " , (case "
            sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.beadaldiv in (50, 51) then m.beadaldiv "
            sqlStr = sqlStr + "     when IsNull(r.userlevel, r2.userlevel) = 0 and m.userid = '' then 99 "
            sqlStr = sqlStr + "     else IsNull(IsNull(r.userlevel, r2.userlevel), '-1') end) "
		end if

		''union all 사용
		tmpSqlStr = sqlUnionStr
		''tmpSqlStr = sqlStr

		if tmpSqlStr="" or isnull(tmpSqlStr) then exit function		' 값이 없을경우 오류내지 않고 팅겨낸다.

		sqlStr = " select top 200 "
		sqlStr = sqlStr + " yyyymm "
		sqlStr = sqlStr + " , sitename "
		sqlStr = sqlStr + " , [db_datamart].[dbo].[fn_GetSellChannelName](T.beadaldiv) as sellChannelName "
		sqlStr = sqlStr + " , sum(orgOrderCnt) as orgOrderCnt "
		sqlStr = sqlStr + " , sum(cancelOrderCnt) as cancelOrderCnt "
		sqlStr = sqlStr + " , sum(returnOrderCnt) as returnOrderCnt "
		sqlStr = sqlStr + " , sum(orgTotalPrice) as orgTotalPrice "
		sqlStr = sqlStr + " , sum(subtotalpriceCouponNotApplied) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + " , sum(totalsum) as totalsum "
		sqlStr = sqlStr + " , sum(totalReducedPrice) as totalReducedPrice "
		sqlStr = sqlStr + " , sum(totalReducedPriceBeasongPay) as totalReducedPriceBeasongPay "
		sqlStr = sqlStr + " , sum(totalItemCouponDiscount) as totalItemCouponDiscount "
		sqlStr = sqlStr + " , sum(totalBonusCouponDiscount) as totalBonusCouponDiscount "
		sqlStr = sqlStr + " , sum(totalPriceBonusCouponDiscount) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , sum(totalBeasongBonusCouponDiscount) as totalBeasongBonusCouponDiscount "
		sqlStr = sqlStr + " , sum(allatdiscountprice) as allatdiscountprice "
		sqlStr = sqlStr + " , sum(totalMaechulPrice) as totalMaechulPrice "
		sqlStr = sqlStr + " , sum(totalMaechulVatPrice) as totalMaechulVatPrice "
		sqlStr = sqlStr + " , sum(mileTotalPrice) as mileTotalPrice "
		sqlStr = sqlStr + " , sum(giftTotalPrice) as giftTotalPrice "
		sqlStr = sqlStr + " , sum(depositTotalPrice) as depositTotalPrice "
		sqlStr = sqlStr + " , sum(totalBuycash) as totalBuycash "
		sqlStr = sqlStr + " , sum(totalBuycashVAT) as totalBuycashVAT "
		sqlStr = sqlStr + " , sum(totalBuycashCouponNotApplied) as totalBuycashCouponNotApplied "
		sqlStr = sqlStr + " , sum(totalUpcheJungsanCash) as totalUpcheJungsanCash "
		sqlStr = sqlStr + " , sum(totalUpcheJungsanCashVAT) as totalUpcheJungsanCashVAT "
		sqlStr = sqlStr + " , sum(totalMileage) as totalMileage "
		if (FRectShowLevel = "Y") then
			sqlStr = sqlStr + " , IsNull(T.userlevel, '-1') as userlevel "
		end if
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "

		sqlStr = sqlStr + tmpSqlStr

		''response.write tmpSqlStr & "<Br>"
		''response.end

		sqlStr = sqlStr + " 	) T "

		if FRectExcTPL <> "" then
			sqlStr = sqlStr + " where T.sitename not in (select distinct id as sitename from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
		end If

		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " yyyymm "
		sqlStr = sqlStr + " , sitename "
		sqlStr = sqlStr + " , [db_datamart].[dbo].[fn_GetSellChannelName](beadaldiv) "
		if (FRectShowLevel = "Y") then
			sqlStr = sqlStr + " , IsNull(T.userlevel, '-1') "
		end if
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " yyyymm "
        sqlStr = sqlStr + " , sitename "
		sqlStr = sqlStr + " , sum(orgOrderCnt) desc "

		''response.write sqlstr & "<Br>"
		''response.end
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulLogSUMItem

				FItemList(i).Fyyyymm = db3_rsget("yyyymm")
				FItemList(i).Fsitename = db3_rsget("sitename")
				FItemList(i).FsellChannelName = db3_rsget("sellChannelName")

				FItemList(i).ForgOrderCnt = db3_rsget("orgOrderCnt")
				FItemList(i).FcancelOrderCnt = db3_rsget("cancelOrderCnt")
				FItemList(i).FreturnOrderCnt = db3_rsget("returnOrderCnt")

				FItemList(i).ForgTotalPrice = db3_rsget("orgTotalPrice")
				FItemList(i).FsubtotalpriceCouponNotApplied = db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).Ftotalsum = db3_rsget("totalsum")
				FItemList(i).FtotalReducedPrice = db3_rsget("totalReducedPrice")
				''FItemList(i).FanbunPriceDetailSUM = db3_rsget("anbunPriceDetailSUM")
				FItemList(i).FtotalMaechulPrice = db3_rsget("totalMaechulPrice")
				FItemList(i).FtotalBuycash = db3_rsget("totalBuycash")
				FItemList(i).FtotalUpcheJungsanCash = db3_rsget("totalUpcheJungsanCash")
				FItemList(i).FtotalMileage = db3_rsget("totalMileage")
				FItemList(i).FtotalBonusCouponDiscount = db3_rsget("totalBonusCouponDiscount")
				FItemList(i).FtotalBeasongBonusCouponDiscount = db3_rsget("totalBeasongBonusCouponDiscount")
				FItemList(i).FtotalPriceBonusCouponDiscount = db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).Fallatdiscountprice = db3_rsget("allatdiscountprice")

				FItemList(i).FmileTotalPrice = db3_rsget("mileTotalPrice")
				FItemList(i).FgiftTotalPrice = db3_rsget("giftTotalPrice")
				FItemList(i).FdepositTotalPrice = db3_rsget("depositTotalPrice")

                if (FRectShowLevel = "Y") then
                    FItemList(i).Fuserlevel = db3_rsget("userlevel")
                end if

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end function

	'//admin/maechul/maechul_detail_log.asp
	public function GetMaechulDetailLog()
	    dim i,sqlStr, addSqlStr, indexmSqlStr, indexdSqlStr
		Dim tmpOrderSerial

		addSqlStr = ""

		if FRectDategbn="ActDate" then
			indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_actDate))"
		elseif FRectDategbn="chulgoDate" then
			indexdSqlStr = indexdSqlStr + " with (NOLOCK,index(IX_tbl_order_detail_log_beasongdate))"
		elseif FRectDategbn="jFixedDt" then
			indexdSqlStr = indexdSqlStr + " with (NOLOCK)"
		else
			indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_ipkumdate))"
		end if

		if (application("Svr_Info")="Dev") then indexmSqlStr=" "
		if (application("Svr_Info")="Dev") then indexdSqlStr=" "

		if FRectOrgPayStartDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectOrgPayStartDate) + "'"
		end if
		if FRectOrgPayEndDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectOrgPayEndDate) + "'"
		end if

		if FRectActDateStartDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectActDateStartDate) + "'"
		end if
		if FRectActDateEndDate <> "" then
			addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectActDateEndDate) + "'"
		end if

		if FRectChulgoDateStartDate <> "" then
			addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectChulgoDateStartDate) + "'"
		end if
		if FRectChulgoDateEndDate <> "" then
			addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectChulgoDateEndDate) + "'"
		end if

		if FRectjFixedDtStartDate <> "" then
			addSqlStr = addSqlStr + " and d.DTLjFixedDt>='" + CStr(FRectjFixedDtStartDate) + "'"
		end if
		if FRectjFixedDtEndDate <> "" then
			addSqlStr = addSqlStr + " and d.DTLjFixedDt<'" + CStr(FRectjFixedDtEndDate) + "'"
		end if

		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if
		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if (FRectMichulgoOnly = "Y") then
			''addSqlStr = addSqlStr + " and d.beasongdate is NULL "
			''addSqlStr = addSqlStr + " and (select IsNull(sum(itemno), 0) from db_datamart.dbo.tbl_order_detail_log s with (nolock) where d.orderserial = s.orderserial and d.itemid = s.itemid and d.itemoption = s.itemoption and s.beasongdate is NULL) <> 0 "
			''addSqlStr = addSqlStr + " and d.canceldate is NULL "

            addSqlStr = addSqlStr + " and d.beasongdate is NULL "
            ''addSqlStr = addSqlStr + " and m.actDivCode not in ('M', 'MM') "
		end if
        if (FRectMiJungsanOnly = "Y") then
            addSqlStr = addSqlStr + " and d.DTLjFixedDt is NULL "
            ''addSqlStr = addSqlStr + " and m.actDivCode not in ('M', 'MM') "
		end if
		if FRectmwdiv_beasongdiv="M" or FRectmwdiv_beasongdiv="W" or FRectmwdiv_beasongdiv="U" or FRectmwdiv_beasongdiv="R" then
			addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + FRectmwdiv_beasongdiv + "'"
		elseif FRectmwdiv_beasongdiv="TT" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
		elseif FRectmwdiv_beasongdiv="UU" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
		elseif (Len(FRectmwdiv_beasongdiv) = 4) then
			addSqlStr = addSqlStr + " and d.omwdiv='" + FRectmwdiv_beasongdiv + "' "
		end if
		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if

		if (FRectExcTPL <> "") and (application("Svr_Info")<>"Dev") then
			addSqlStr = addSqlStr + " and m.sitename not in (select id as sitename from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
		end if

		if FRectExcZeroPrice <> "" then
			addSqlStr = addSqlStr + " and (d.orgitemcost <> 0 or d.anbunAppliedPriceDetailSUM <> 0) "
		end if

		If FRectExc6month <> "" Then
			tmpOrderSerial = Right(Replace(DateSerial(Year(Now()), Month(Now()) - 3, 1), "-", ""), 6) & "00000"
			addSqlStr = addSqlStr + " and m.orderserial >= '" & tmpOrderSerial & "' "
			addSqlStr = addSqlStr + " and m.suborderserial >= 0 "
		End If

		''addSqlStr = addSqlStr + " and m.orderserial >= '13060100000' "
		''addSqlStr = addSqlStr + " and d.orderserial >= '13060100000' "

		sqlStr = "select IsNull(sum(d.orgitemcost * d.itemno), 0) as orgTotalPrice  "
		sqlStr = sqlStr + " , IsNull(sum(d.itemcostCouponNotApplied * d.itemno), 0) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(sum(d.itemcost * d.itemno), 0) as totalsum "
		sqlStr = sqlStr + " , IsNull(sum(d.reducedPrice * d.itemno), 0) as totalReducedPrice "
		sqlStr = sqlStr + " , IsNull(sum(d.anbunPriceDetailSUM), 0) as anbunPriceDetailSUM "
		sqlStr = sqlStr + " , IsNull(sum(d.anbunAppliedPriceDetailSUM), 0) as totalMaechulPrice "
		sqlStr = sqlStr + " , IsNull(sum(d.buycash * d.itemno), 0) as totalBuycash "
		sqlStr = sqlStr + " , IsNull(sum(d.upcheJungsanCash * d.itemno), 0) as totalUpcheJungsanCash "
		sqlStr = sqlStr + " , IsNull(sum(d.mileage * d.itemno), 0) as totalMileage "
		sqlStr = sqlStr + " , IsNull(sum((d.itemcost - d.reducedPrice - IsNull(d.allAtDiscount, 0)) * d.itemno), 0) as totalBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(sum((case when d.itemid = 0 then d.itemcost - d.reducedPrice else 0 end) * d.itemno), 0) as totalBeasongBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(sum(d.anbunCouponPriceDetailSUM), 0) as totalPriceBonusCouponDiscount "
		sqlStr = sqlStr + " , IsNull(sum(IsNull(d.allAtDiscount, 0) * d.itemno), 0) as allatdiscountprice "
		sqlStr = sqlStr + " , IsNull(sum((case "
		sqlStr = sqlStr + " 	when d.omwdiv in ('M', 'B031') then s.avgipgoPrice*d.itemno "
		sqlStr = sqlStr + "     else 0 end)),0) as avgipgoPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr + "		join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			1 = 1 "
		sqlStr = sqlStr + "			and m.orderserial = d.orderserial "
		sqlStr = sqlStr + "			and m.suborderserial = d.suborderserial "
		sqlStr = sqlStr + "		Left Join tendb.db_summary.dbo.tbl_monthly_accumulated_logisstock_summary as s with(noLock) "
		sqlStr = sqlStr + "		on s.yyyymm=convert(varchar(7),m.actDate,21) "
		sqlStr = sqlStr + "			and s.itemgubun=d.itemgubun "
		sqlStr = sqlStr + "			and s.itemid=d.itemid "
		sqlStr = sqlStr + "			and s.itemoption=d.itemoption "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + addSqlStr

        if (session("ssBctId") = "skyer9") then
            response.write sqlStr
        end if

		'response.end
		if (FRectShowStatistic = "Y") then
		    ''db3_dbget.CommandTimeout = 60  ''2016/01/07 (기본 30초)
			db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
				Set FOneItem = new CMaechulLogSUMItem

				FOneItem.ForgTotalPrice = db3_rsget("orgTotalPrice")
				FOneItem.FsubtotalpriceCouponNotApplied = db3_rsget("subtotalpriceCouponNotApplied")
				FOneItem.Ftotalsum = db3_rsget("totalsum")
				FOneItem.FtotalReducedPrice = db3_rsget("totalReducedPrice")
				''FOneItem.FanbunPriceDetailSUM = db3_rsget("anbunPriceDetailSUM")
				FOneItem.FtotalMaechulPrice = db3_rsget("totalMaechulPrice")
				FOneItem.FtotalBuycash = db3_rsget("totalBuycash")
				FOneItem.FtotalUpcheJungsanCash = db3_rsget("totalUpcheJungsanCash")
				FOneItem.FtotalMileage = db3_rsget("totalMileage")
				FOneItem.FtotalBonusCouponDiscount = db3_rsget("totalBonusCouponDiscount")
				FOneItem.FtotalBeasongBonusCouponDiscount = db3_rsget("totalBeasongBonusCouponDiscount")
				FOneItem.FtotalPriceBonusCouponDiscount = db3_rsget("totalPriceBonusCouponDiscount")
				FOneItem.Fallatdiscountprice = db3_rsget("allatdiscountprice")
				FOneItem.FavgipgoPrice = db3_rsget("avgipgoPrice")
			db3_rsget.Close



		end if

        if (FRectshowOnlyStatistic = "Y") then
            exit function ''합계만표시시 그만 쿼리..
        end if

		sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and m.suborderserial = d.suborderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.actDivCode, m.actDate, m.jumundiv, m.accountdiv, m.sitename, m.ipkumdate, m.regdate, IsNull(m.targetGbn, 'ON') as targetGbn, d.*, IsNull(d.allAtDiscount, 0) as allAtDiscount, IsNull(d.orgitemcost, 0) as orgitemcost, IsNull(d.itemcostCouponNotApplied, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " , IsNull((case "
		sqlStr = sqlStr + " 	when d.omwdiv in ('M', 'B031') then s.avgipgoPrice*d.itemno "
		sqlStr = sqlStr + "     else 0 end),0) as avgipgoPrice "
		sqlStr = sqlStr + " , (select top 1 linkorderserial from [db_order].[dbo].[tbl_order_master] o where o.orderserial = m.orderserial) as orgorderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and m.suborderserial = d.suborderserial "
		sqlStr = sqlStr + "		Left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary as s with(noLock) "
		sqlStr = sqlStr + "		on s.yyyymm=convert(varchar(7),m.actDate,21) "
		sqlStr = sqlStr + "			and s.itemgubun=d.itemgubun "
		sqlStr = sqlStr + "			and s.itemid=d.itemid "
		sqlStr = sqlStr + "			and s.itemoption=d.itemoption "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
    	sqlStr = sqlStr + " order by m.actDate desc"

		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).FtargetGbn		= db3_rsget("targetGbn")
				FItemList(i).Forderserial		= db3_rsget("orderserial")
				FItemList(i).Fsuborderserial	= db3_rsget("suborderserial")

				FItemList(i).Forgorderserial	= db3_rsget("orgorderserial")

				FItemList(i).FactDivCode		= db3_rsget("actDivCode")
				FItemList(i).FactDate			= db3_rsget("actDate")
				FItemList(i).Fjumundiv			= db3_rsget("jumundiv")
				FItemList(i).Faccountdiv		= db3_rsget("accountdiv")

				FItemList(i).Fitemid			= db3_rsget("itemid")
				FItemList(i).Fitemoption		= db3_rsget("itemoption")
				FItemList(i).Fitemno			= db3_rsget("itemno")
				FItemList(i).Fitemname			= db3_rsget("itemname")
				FItemList(i).Fitemoptionname	= db3_rsget("itemoptionname")
				FItemList(i).Fmakerid			= db3_rsget("makerid")

				FItemList(i).ForgTotalPrice						= db3_rsget("orgitemcost") * db3_rsget("itemno")
				FItemList(i).FsubtotalpriceCouponNotApplied		= db3_rsget("itemcostCouponNotApplied") * db3_rsget("itemno")
				FItemList(i).Ftotalsum							= db3_rsget("itemcost") * db3_rsget("itemno")
				FItemList(i).FtotalReducedPrice					= db3_rsget("reducedPrice") * db3_rsget("itemno")

				FItemList(i).FanbunPriceDetailSUM				= db3_rsget("anbunPriceDetailSUM")
				FItemList(i).FtotalMaechulPrice					= db3_rsget("anbunAppliedPriceDetailSUM")

				FItemList(i).FtotalBuycash						= db3_rsget("buycash") * db3_rsget("itemno")
				FItemList(i).FtotalUpcheJungsanCash				= db3_rsget("upcheJungsanCash") * db3_rsget("itemno")

				FItemList(i).FtotalMileage						= db3_rsget("mileage") * db3_rsget("itemno")
				FItemList(i).Fipkumdate							= db3_rsget("ipkumdate")
				FItemList(i).Fsitename							= db3_rsget("sitename")
				''FItemList(i).Frdsite							= db3_rsget("rdsite")
				FItemList(i).Fregdate							= db3_rsget("regdate")

				FItemList(i).FmileTotalPrice					= 0
				FItemList(i).FgiftTotalPrice					= 0
				FItemList(i).FdepositTotalPrice					= 0

				FItemList(i).Fvatinclude						= db3_rsget("vatinclude")
				FItemList(i).Fomwdiv							= db3_rsget("omwdiv")

				FItemList(i).FtotalBonusCouponDiscount			= (db3_rsget("itemcost") - db3_rsget("reducedPrice") - db3_rsget("allAtDiscount")) * db3_rsget("itemno")

				if (db3_rsget("itemid") = 0) then
					FItemList(i).FtotalBeasongBonusCouponDiscount = (db3_rsget("itemcost") - db3_rsget("reducedPrice")) * db3_rsget("itemno")
				else
					FItemList(i).FtotalBeasongBonusCouponDiscount = 0
				end if

				FItemList(i).FtotalPriceBonusCouponDiscount 	= db3_rsget("anbunCouponPriceDetailSUM")
				FItemList(i).Fallatdiscountprice 				= db3_rsget("allAtDiscount") * db3_rsget("itemno")

				FItemList(i).Fbeasongdate						= db3_rsget("beasongdate")
				FItemList(i).FDTLjFixedDt						= db3_rsget("DTLjFixedDt")
				FItemList(i).Fcanceldate						= db3_rsget("canceldate")
				FItemList(i).FtotalMileage						= db3_rsget("mileage") * db3_rsget("itemno")
				FItemList(i).FavgipgoPrice						= db3_rsget("avgipgoPrice")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	public function GetMaechulPaymentLog()
	    dim i,sqlStr, addSqlStr

		addSqlStr = ""

		if FRectStartdate <> "" and FRectEndDate <> "" then
			Select Case FRectDateGubun
				Case "paydate"
					addSqlStr = addSqlStr + " and m.paydate>='" + CStr(FRectStartdate) + "'"
					addSqlStr = addSqlStr + " and m.paydate<'" + CStr(FRectEndDate) + "'"
				Case "maeipdate"
					addSqlStr = addSqlStr + " and m.maeipdate>='" + CStr(FRectStartdate) + "'"
					addSqlStr = addSqlStr + " and m.maeipdate<'" + CStr(FRectEndDate) + "'"
				Case "mayipkumdate"
					addSqlStr = addSqlStr + " and m.mayipkumdate>='" + CStr(FRectStartdate) + "'"
					addSqlStr = addSqlStr + " and m.mayipkumdate<'" + CStr(FRectEndDate) + "'"
				Case Else
					addSqlStr = addSqlStr + " and m.payReqDate>='" + CStr(FRectStartdate) + "'"
					addSqlStr = addSqlStr + " and m.payReqDate<'" + CStr(FRectEndDate) + "'"
			End Select

            if FRectIncActPayMonthDiff <> "" then
                addSqlStr = addSqlStr + " and DateDiff(month, m.payReqDate, m.paydate) <> 0 "
            end if
		end if

		'// 모지??, skyer9, 2015-11-11
		if FRectStartdate <> "" then
			''addSqlStr = addSqlStr + " and m.payReqDate>='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate <> "" then
			''addSqlStr = addSqlStr + " and m.payReqDate<'" + CStr(FRectEndDate) + "'"
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField = "orderserial") then
				''addSqlStr = addSqlStr + " and ((m.orderserial = '" + CStr(FRectSearchText) + "') or (IsNull(m.orgorderserial, m.orderserial) = '" + CStr(FRectSearchText) + "')) "
				addSqlStr = addSqlStr + " and ((m.orderserial = '" + CStr(FRectSearchText) + "') or (m.orgorderserial = '" + CStr(FRectSearchText) + "')) " ''2017/02/08 수정.orgorderserial 인덱스추가
			else
				addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectPayDivCode <> "") then
			if (FRectPayDivCode = "etc") then
				addSqlStr = addSqlStr + " and m.payDivCode = '' "
			else
				addSqlStr = addSqlStr + " and m.payDivCode = '" + CStr(FRectPayDivCode) + "' "
			end if
		end if

		if (FRectTargetGbn <> "") then
			addSqlStr = addSqlStr + " and m.targetGbn = '" + CStr(FRectTargetGbn) + "' "
		end if

		if (FRectShowOnlyPriceNotMatch = "Y") then
			addSqlStr = addSqlStr + " and m.payReqPrice <> IsNull(m.realPayPrice, 0) "
		end if

		if (FRectExcNoPay = "Y") then
			addSqlStr = addSqlStr + " and not (m.payReqPrice = 0 and IsNull(m.realPayPrice, 0) = 0) "
			addSqlStr = addSqlStr + " and (m.PGgubun <> 'nopayment') "
		end if

		if (FRectExcNoReqPay = "Y") then
			addSqlStr = addSqlStr + " and not (m.payReqPrice = 0 and IsNull(m.realPayPrice, 0) = 0) "
			addSqlStr = addSqlStr + " and (m.PGgubun <> 'XXX') "
		end if

		if (FRectExcHP = "Y") then
			addSqlStr = addSqlStr + " and IsNull(l.accountdiv, '') not in ('400') "
		end if

		if (FRectExcGift = "Y") then
			'// 승인내역이 늦게 온다.
			addSqlStr = addSqlStr + " and IsNull(l.accountdiv, '') not in ('550', '560') "
		end if

		if (FRectMatchState <> "") then
			if (FRectMatchState = "Y") then
				addSqlStr = addSqlStr + " and IsNull(m.matchMethod, 'X') <> 'X' "
			else
				addSqlStr = addSqlStr + " and IsNull(m.matchMethod, 'X') = '" + CStr(FRectMatchState) + "' "
			end if
		end if

		if (FRectPGgubun <> "") then
			addSqlStr = addSqlStr + " and m.PGgubun = '" + CStr(FRectPGgubun) + "' "
		end if

		if (FRectPGuserid <> "") then
			addSqlStr = addSqlStr + " and m.PGuserid = '" + CStr(FRectPGuserid) + "' "
		end if

		sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg, IsNull(sum(m.payReqPrice),0) as payReqPrice, IsNull(sum(m.realPayPrice),0) as realPayPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log m with (nolock) "
		sqlStr = sqlStr + " 	left join db_datamart.dbo.tbl_order_master_log l with (nolock) "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.orderserial = l.orderserial "
		sqlStr = sqlStr + " 	and m.suborderserial = l.suborderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		response.write sqlstr & "<Br>"
        ''response.end
		db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")

            FTotalPayReqPrice = db3_rsget("payReqPrice")
            FTotalRealPayPrice = db3_rsget("realPayPrice")
		db3_rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, IsNull(l.accountdiv, '') as mayPayMethod, IsNull(l.actDivCode, '') as actDivCode "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log m with (nolock) "
		sqlStr = sqlStr + " 	left join db_datamart.dbo.tbl_order_master_log l with (nolock) "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.orderserial = l.orderserial "
		sqlStr = sqlStr + " 	and m.suborderserial = l.suborderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by m.payReqDate desc, m.orderserial, m.suborderserial desc, m.payDivCode"

		''response.write sqlStr & "<Br>"
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulPaymentLogItem

				FItemList(i).FtargetGbn			= db3_rsget("targetGbn")
				FItemList(i).Forderserial		= db3_rsget("orderserial")
				FItemList(i).Fsuborderserial	= db3_rsget("suborderserial")
				FItemList(i).Forgorderserial	= db3_rsget("orgorderserial")
				FItemList(i).Fchgorderserial	= db3_rsget("chgorderserial")
				FItemList(i).FpayDivCode		= db3_rsget("payDivCode")
				FItemList(i).FPGgubun			= db3_rsget("PGgubun")
				FItemList(i).FPGuserid			= db3_rsget("PGuserid")
				FItemList(i).FPGkey				= db3_rsget("PGkey")
				FItemList(i).FPGCSkey			= db3_rsget("PGCSkey")
				FItemList(i).FpayReqPrice		= db3_rsget("payReqPrice")
				FItemList(i).FrealPayPrice		= db3_rsget("realPayPrice")
				FItemList(i).FpayReqDate		= db3_rsget("payReqDate")
				FItemList(i).FpayDate			= db3_rsget("payDate")
				FItemList(i).FmaeipDate			= db3_rsget("maeipDate")
				FItemList(i).FmayIpkumDate		= db3_rsget("mayIpkumDate")
				FItemList(i).Fregdate			= db3_rsget("regdate")

				FItemList(i).FmayPayMethod		= db3_rsget("mayPayMethod")
				FItemList(i).FactDivCode		= db3_rsget("actDivCode")

				FItemList(i).FmatchMethod		= db3_rsget("matchMethod")

				FItemList(i).FcommPrice         = db3_rsget("commPrice")
				'FItemList(i).FcommVatPrice      = db3_rsget("commVatPrice")
                FItemList(i).FjungsanPrice      = db3_rsget("jungsanPrice")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	public function GetMaechulPaymentLogCheck()
	    dim i,sqlStr, addSqlStr

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " IsNull(o.targetGbn, p.targetGbn) as targetGbn, IsNull(o.actdate, p.payreqdate) as actdate, IsNull(o.orderserial, p.orderserial) as orderserial, IsNull(o.totalOrderMaechulPrice, 0) as totalOrderMaechulPrice, IsNull(p.totalpayreqPrice, 0) as totalpayreqPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + " 		select "
		sqlStr = sqlStr + " 			IsNull(targetGbn, 'ON') as targetGbn, convert(varchar(10),actdate,121) as actdate, IsNull(sum(totalMaechulPrice),0) as totalOrderMaechulPrice "

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = sqlStr + " 			, orderserial "
		else
			sqlStr = sqlStr + " 			, '' as orderserial "
		end if

		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			db_datamart.dbo.tbl_order_master_log with (nolock) "
		sqlStr = sqlStr + " 		where "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and actdate >= '" + CStr(FRectStartdate) + "' "
		sqlStr = sqlStr + " 			and actdate < '" + CStr(FRectEndDate) + "' "
		sqlStr = sqlStr + " 			and IsNull(targetGbn, 'ON') = '" + CStr(FRectTargetGbn) + "' "
		sqlStr = sqlStr + " 			and IsNull(sitename, '') in ('10x10', '10x10_cs') "
		sqlStr = sqlStr + " 		group by "
		sqlStr = sqlStr + " 			IsNull(targetGbn, 'ON'), convert(varchar(10),actdate,121) "

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = sqlStr + " 			, orderserial "
		end if

		sqlStr = sqlStr + " 	) o "
		sqlStr = sqlStr + " 	FULL OUTER join ( "
		sqlStr = sqlStr + " 		select "
		sqlStr = sqlStr + " 			IsNull(targetGbn, 'ON') as targetGbn, convert(varchar(10),payreqdate,121) as payreqdate, IsNull(sum(payreqPrice),0) as totalpayreqPrice "

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = sqlStr + " 			, orderserial "
		else
			sqlStr = sqlStr + " 			, '' as orderserial "
		end if

		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			db_datamart.dbo.tbl_order_payment_log with (nolock) "
		sqlStr = sqlStr + " 		where "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and payreqdate >= '" + CStr(FRectStartdate) + "' "
		sqlStr = sqlStr + " 			and payreqdate < '" + CStr(FRectEndDate) + "' "
		sqlStr = sqlStr + " 			and IsNull(targetGbn, 'ON') = '" + CStr(FRectTargetGbn) + "' "
		sqlStr = sqlStr + " 		group by "
		sqlStr = sqlStr + " 			IsNull(targetGbn, 'ON'), convert(varchar(10),payreqdate,121) "

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = sqlStr + " 			, orderserial "
		end if

		sqlStr = sqlStr + " 	) p "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and o.targetGbn = p.targetGbn "
		sqlStr = sqlStr + " 		and o.actdate = p.payreqdate "
		sqlStr = sqlStr + " 		and o.orderserial = p.orderserial "

		if (FRectChkGrpByOrderserial = "Y") Then
			sqlStr = sqlStr + " where 1 = 1 "
			sqlStr = sqlStr + " 			and IsNull(o.totalOrderMaechulPrice,0) <> IsNull(p.totalpayreqPrice, 0) "
		end if

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	convert(varchar(10),o.actdate,121) "

		if (FRectChkGrpByOrderserial = "Y") then
			sqlStr = sqlStr + " 			, o.orderserial "
		end if

        if (FRectChkGrpByOrderserial = "Y") then
            sqlStr = " exec db_datamart.dbo.[usp_TEN_OrderLog_OrderAndPay_Diff] '" + CStr(FRectStartdate) + "','" + CStr(FRectEndDate) + "','" + CStr(FRectTargetGbn) + "' "
        end if

		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		''db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/05
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulPaymentLogCheckItem

				FItemList(i).FtargetGbn					= db3_rsget("targetGbn")
				FItemList(i).Factdate					= db3_rsget("actdate")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).FtotalOrderMaechulPrice	= db3_rsget("totalOrderMaechulPrice")
				FItemList(i).FtotalpayreqPrice			= db3_rsget("totalpayreqPrice")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

    public function GetMaechulMasterDetailLogCheck()
        dim i,sqlStr, addSqlStr

        sqlStr = " EXEC [db_datamart].[dbo].[usp_Ten_OrderLogMasterDetailDiff_List_ON] " + CStr(FPageSize*FCurrPage) + ", '" & FRectStartdate & "', '" & FRectEndDate & "' "

		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/05

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulPaymentLogCheckItem

				FItemList(i).Factdate					= db3_rsget("yyyymmdd")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).FtotalOrderMaechulPrice	= db3_rsget("totalReducedPrice")
				FItemList(i).FtotalpayreqPrice			= db3_rsget("totalReducedPriceDetail")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end function

    '// 승인내역 기준 매출로그 검토
	public function GetPaymentMaechulLogCheck()
	    dim i,sqlStr, addSqlStr

        if (FRectDateGubun = "log") then
		    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " 'ON' as targetGbn, T1.yyyymmdd, '' as orderserial, IsNull(T2.appPrice,0) appPrice, IsNull(T1.realPayPrice,0) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + " from " & vbCrLf
		    sqlStr = sqlStr + "     ( " & vbCrLf
		    sqlStr = sqlStr + "         select convert(varchar(10), payDate, 121) yyyymmdd, sum(realPayPrice) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_datamart].[dbo].[tbl_order_payment_log] p with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and payDate >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and payDate < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             convert(varchar(10), payDate, 121) " & vbCrLf
		    sqlStr = sqlStr + "     ) T1 " & vbCrLf
		    sqlStr = sqlStr + "     left join ( " & vbCrLf
		    sqlStr = sqlStr + "         select convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) yyyymmdd, sum(appPrice) appPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_order].[dbo].[tbl_onlineApp_log] l with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) " & vbCrLf
		    sqlStr = sqlStr + "     ) T2 " & vbCrLf
		    sqlStr = sqlStr + "     on " & vbCrLf
		    sqlStr = sqlStr + "         T1.yyyymmdd = T2.yyyymmdd " & vbCrLf
		    sqlStr = sqlStr + " where " & vbCrLf
		    sqlStr = sqlStr + "     1 = 1 " & vbCrLf
            ''sqlStr = sqlStr + "     and T1.realPayPrice <> T2.appPrice " & vbCrLf
		    sqlStr = sqlStr + " order by " & vbCrLf
		    sqlStr = sqlStr + "     T1.yyyymmdd " & vbCrLf
        else
		    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " 'ON' as targetGbn, T1.yyyymmdd, '' as orderserial, IsNull(T1.appPrice,0) appPrice, IsNull(T2.realPayPrice,0) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + " from " & vbCrLf
		    sqlStr = sqlStr + "     ( " & vbCrLf
		    sqlStr = sqlStr + "         select convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) yyyymmdd, sum(appPrice) appPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_order].[dbo].[tbl_onlineApp_log] l with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) " & vbCrLf
		    sqlStr = sqlStr + "     ) T1 " & vbCrLf
		    sqlStr = sqlStr + "     join ( " & vbCrLf
		    sqlStr = sqlStr + "         select convert(varchar(10), payDate, 121) yyyymmdd, sum(realPayPrice) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_datamart].[dbo].[tbl_order_payment_log] p with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and payDate >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and payDate < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             convert(varchar(10), payDate, 121) " & vbCrLf
		    sqlStr = sqlStr + "     ) T2 " & vbCrLf
		    sqlStr = sqlStr + "     on " & vbCrLf
		    sqlStr = sqlStr + "         T1.yyyymmdd = T2.yyyymmdd " & vbCrLf
		    sqlStr = sqlStr + " where " & vbCrLf
            sqlStr = sqlStr + "     1 = 1 " & vbCrLf
		    ''sqlStr = sqlStr + "     and T2.realPayPrice <> T1.appPrice " & vbCrLf
		    sqlStr = sqlStr + " order by " & vbCrLf
		    sqlStr = sqlStr + "     T1.yyyymmdd " & vbCrLf
        end if

		''response.write sqlStr & "<Br>"
		''response.end
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/05

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulPaymentLogCheckItem

				FItemList(i).FtargetGbn					= db3_rsget("targetGbn")
				FItemList(i).Factdate					= db3_rsget("yyyymmdd")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).FappPrice					= db3_rsget("appPrice")
				FItemList(i).FrealPayPrice				= db3_rsget("realPayPrice")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	public function GetPaymentMaechulLogOneDayCheck()
	    dim i,sqlStr, addSqlStr

        if (FRectDateGubun = "log") then
		    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " 'ON' as targetGbn, T1.yyyymmdd, T1.orderserial, IsNull(T2.appPrice,0) as appPrice, IsNull(T1.realPayPrice, 0) as realPayPrice " & vbCrLf
		    sqlStr = sqlStr + " from " & vbCrLf
		    sqlStr = sqlStr + "     ( " & vbCrLf
		    sqlStr = sqlStr + "         select pgkey, pgcskey, convert(varchar(10), payDate, 121) yyyymmdd, orderserial, sum(realPayPrice) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_datamart].[dbo].[tbl_order_payment_log] p with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and payDate >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and payDate < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             pgkey, pgcskey, convert(varchar(10), payDate, 121), orderserial " & vbCrLf
		    sqlStr = sqlStr + "     ) T1 " & vbCrLf
		    sqlStr = sqlStr + "     left join ( " & vbCrLf
		    sqlStr = sqlStr + "         select pgkey, pgcskey, convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) yyyymmdd, orderserial, sum(appPrice) appPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_order].[dbo].[tbl_onlineApp_log] l with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "             and pgkey not like '%[_]1' " & vbCrLf
		    sqlStr = sqlStr + "             and pgkey not like '%[_]2' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             pgkey, pgcskey, convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121), orderserial " & vbCrLf
		    sqlStr = sqlStr + "     ) T2 " & vbCrLf
		    sqlStr = sqlStr + "     on " & vbCrLf
		    sqlStr = sqlStr + "         T1.yyyymmdd = T2.yyyymmdd and T1.pgkey = T2.pgkey and T1.pgcskey = T2.pgcskey " & vbCrLf
		    sqlStr = sqlStr + " where " & vbCrLf
		    sqlStr = sqlStr + "     IsNull(T2.appPrice,0) <> IsNull(T1.realPayPrice,0) " & vbCrLf
		    sqlStr = sqlStr + " order by " & vbCrLf
		    sqlStr = sqlStr + "     T1.yyyymmdd, T1.orderserial " & vbCrLf
        else
		    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " 'ON' as targetGbn, T1.yyyymmdd, T1.orderserial, T1.appPrice, IsNull(T2.realPayPrice, 0) as realPayPrice " & vbCrLf
		    sqlStr = sqlStr + " from " & vbCrLf
		    sqlStr = sqlStr + "     ( " & vbCrLf
		    sqlStr = sqlStr + "         select pgkey, pgcskey, convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121) yyyymmdd, orderserial, sum(appPrice) appPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_order].[dbo].[tbl_onlineApp_log] l with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "             and pgkey not like '%[_]1' " & vbCrLf
		    sqlStr = sqlStr + "             and pgkey not like '%[_]2' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             pgkey, pgcskey, convert(varchar(10), IsNull(l.cancelDate, l.appDate), 121), orderserial " & vbCrLf
		    sqlStr = sqlStr + "     ) T1 " & vbCrLf
		    sqlStr = sqlStr + "     left join ( " & vbCrLf
		    sqlStr = sqlStr + "         select pgkey, pgcskey, convert(varchar(10), payDate, 121) yyyymmdd, orderserial, sum(realPayPrice) realPayPrice " & vbCrLf
		    sqlStr = sqlStr + "         from " & vbCrLf
		    sqlStr = sqlStr + "         [db_datamart].[dbo].[tbl_order_payment_log] p with (nolock) " & vbCrLf
		    sqlStr = sqlStr + "         where " & vbCrLf
		    sqlStr = sqlStr + "             1 = 1 " & vbCrLf
		    sqlStr = sqlStr + "             and payDate >= '" + CStr(FRectStartdate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and payDate < '" + CStr(FRectEndDate) + "' " & vbCrLf
		    sqlStr = sqlStr + "             and PGgubun = '" & FRectPGGubun & "' " & vbCrLf
		    sqlStr = sqlStr + "         group by " & vbCrLf
		    sqlStr = sqlStr + "             pgkey, pgcskey, convert(varchar(10), payDate, 121), orderserial " & vbCrLf
		    sqlStr = sqlStr + "     ) T2 " & vbCrLf
		    sqlStr = sqlStr + "     on " & vbCrLf
		    sqlStr = sqlStr + "         T1.yyyymmdd = T2.yyyymmdd and T1.pgkey = T2.pgkey and T1.pgcskey = T2.pgcskey " & vbCrLf
		    sqlStr = sqlStr + " where " & vbCrLf
		    sqlStr = sqlStr + "     T1.appPrice <> IsNull(T2.realPayPrice,0) " & vbCrLf
		    sqlStr = sqlStr + " order by " & vbCrLf
		    sqlStr = sqlStr + "     T1.yyyymmdd, T1.orderserial " & vbCrLf
        end if

		''response.write sqlStr & "<Br>"
		''response.end
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/05

		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMaechulPaymentLogCheckItem

				FItemList(i).FtargetGbn					= db3_rsget("targetGbn")
				FItemList(i).Factdate					= db3_rsget("yyyymmdd")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).FappPrice					= db3_rsget("appPrice")
				FItemList(i).FrealPayPrice				= db3_rsget("realPayPrice")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		IF application("Svr_Info")="Dev" THEN
			tendb="tendb."
		end if

		FCurrPage =1
		FPageSize = 20
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
End Class
%>
