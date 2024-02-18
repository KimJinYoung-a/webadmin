<%
'###########################################################
' Description : PG사 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2023.06.01 한용민 수정(편의점결제 추가)
'###########################################################

Class CPGDataItem

	public Fidx
	public FPGgubun
	public FPGkey
	public FappDivCode
	public FappMethod
	public FappDate
	public FcardReaderID
	public FcardPrice
	public FcardAppNo
	public Fshopid
	public FshopJumunMasterIdx
	public Fregdate

	public FcardGubun
	public FcardComp
	public FcardAffiliateNo

	public Fpgmeachuldate
	public Fipkumdate
	public FipkumPrice
	public FcardChargePrice

	public Forderserial
	public ForderCardPrice

	public Flogorderserial
	public Flogsuborderserial

	public FPGCSkey
	public Fsitename
	public FcancelDate
	public FappPrice
	public FcommPrice
	public FcommVatPrice
	public FjungsanPrice
	public Fcsasid
	public FPGuserid

	public FreasonGubun

	public function GetFullLogOrderSerial
		GetFullLogOrderSerial = Flogorderserial + "-" + Format00(3, Flogsuborderserial)
	end function

	public Function GetReasonGubunName
		Select Case FreasonGubun
			Case "001"
				GetReasonGubunName = "선수금(매출)"
			Case "002"
				GetReasonGubunName = "선수금(제휴사 매출)"
			Case "003"
				GetReasonGubunName = "선수금(이니랜탈)"
			Case "020"
				GetReasonGubunName = "선수금(예치금)"
			Case "025"
				GetReasonGubunName = "선수금(예치금환급)"
			Case "030"
				GetReasonGubunName = "선수금(기프트)"
			Case "035"
				GetReasonGubunName = "선수금(기프트환급)"
			Case "040"
				GetReasonGubunName = "CS서비스"
			Case "950"
				GetReasonGubunName = "무통장미확인"
			Case "999"
				GetReasonGubunName = "취소매칭"
			Case "900"
				GetReasonGubunName = "기타"
			Case "901"
				GetReasonGubunName = "핑거스현금매출"
			Case "800"
				GetReasonGubunName = "이자수익"
			Case Else
				GetReasonGubunName = FreasonGubun
		End Select
	end Function

	public function GetAppDivCodeName
		if (FappDivCode = "A") then
			GetAppDivCodeName = "승인"
		elseif (FappDivCode = "C") then
			GetAppDivCodeName = "취소"
		elseif (FappDivCode = "P") then
			GetAppDivCodeName = "전일취소"
		elseif (FappDivCode = "R") then
			GetAppDivCodeName = "부분취소"
		else
			GetAppDivCodeName = FappDivCode
		end if
	end function

	public function GetAppDivCodeColor
		if (FappDivCode = "A") then
			GetAppDivCodeColor = "black"
		elseif (FappDivCode = "C") then
			GetAppDivCodeColor = "red"
		elseif (FappDivCode = "P") then
			GetAppDivCodeColor = "red"
		elseif (FappDivCode = "R") then
			GetAppDivCodeColor = "red"
		else
			GetAppDivCodeColor = "red"
		end if
	end function

	Public function GetAppMethodName()
		if FappMethod = "7" then
			GetAppMethodName = "가상"
		elseif FappMethod = "14" then
			GetAppMethodName = "편의점결제"
		elseif FappMethod = "100" then
			GetAppMethodName = "신용"
		elseif FappMethod = "20" then
			GetAppMethodName = "실시간"
		elseif FappMethod = "30" then
			GetAppMethodName = "포인트"
		elseif FappMethod = "50" then
			GetAppMethodName = "입점몰"
		elseif FappMethod = "80" then
			GetAppMethodName = "All@"
		elseif FappMethod = "90" then
			GetAppMethodName = "상품권"
		elseif FappMethod = "110" then
			GetAppMethodName = "OK캐시백"
		elseif FappMethod = "400" then
			GetAppMethodName = "핸드폰"
		elseif FappMethod = "550" then
			GetAppMethodName = "기프팅"
		elseif FappMethod = "560" then
			GetAppMethodName = "기프티콘"
		elseif FappMethod = "77" then
			GetAppMethodName = "무통장환불"
		elseif FappMethod = "6" then
			GetAppMethodName = "무통장입금"
		elseif FappMethod = "150" then
			GetAppMethodName = "이니랜탈"
		else
			GetAppMethodName = FappMethod
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CPGDataStatisticItem

	public Fshopid
	public Fyyyymmdd
	public FarrSumCardPrice
	public FarrSumCardIpkumPrice
	public FtotSumCardPrice
	public FtotSumCardIpkumPrice
	public FscmTotCardPrice
	public FcardPriceNotMatch

	public FsumCardPrice
	public FsumBankPrice
	public FsumVBankPrice
	public FtotSumPrice
	public FsumCardJungsanPrice
	public FsumBankJungsanPrice
	public FsumVBankJungsanPrice
	public FtotSumJungsanPrice

	public FsumHPPrice
	public FsumHPJungsanPrice
	public FsumGifttingPrice
	public FsumGifticonPrice
	public FsumOKPrice
	public FsumAllAtPrice
	public FsumGifttingJungsanPrice
	public FsumGifticonJungsanPrice
	public FsumOKJungsanPrice
	public FsumAllAtJungsanPrice

	public FsumTenOutBankPrice
	public FsumTenInBankPrice
	public FsumTenOutBankJungsanPrice
	public FsumTenInBankJungsanPrice

	public Fsumteenxteen3Price
	public Fsumteenxteen4Price
	public Fsumteenxteen5Price
	public Fsumteenxteen6Price
	public Fsumteenxteen8Price
	public Fsumteenxteen9Price
	public Fsumteenteen10Price
	public Fsumtenbyten01Price
	public Fsumtenbyten02Price
	public FsumteenxteehaPrice
	public FsumteenxteenrPrice
	public FsumteenteenspPrice
	public FsumteenteenapPrice
	public FsumKCTEN0001mPrice
	public FsumnaverpayPrice
	public FsumnaverpayPoint
	public FsumpaycoPrice
	public FsumbankipkumPrice
	public Fsumbankipkum_10x10Price
	public Fsumbankipkum_fingersPrice
	public FsumbankrefundPrice
	public Fsumbankrefund_10x10Price
	public Fsumbankrefund_fingersPrice
	public Fsum10x10_2Price
	public FsumR5523Price
	public FsummobiliansPrice
	public FsumPGgifticonPrice
	public FsumPGgifttingPrice
	public FsumPGokcashbagPrice
	Public FsumPGtossPrice
	Public FsumPGchaiPrice
	Public FsumPGConvinienspayPrice

	public Fsumteenxteen3JungsanPrice
	public Fsumteenxteen4JungsanPrice
	public Fsumteenxteen5JungsanPrice
	public Fsumteenxteen6JungsanPrice
	public Fsumteenxteen8JungsanPrice
	public Fsumteenxteen9JungsanPrice
	public Fsumteenteen10JungsanPrice
	public Fsumtenbyten01JungsanPrice
	public Fsumtenbyten02JungsanPrice
	public FsumteenxteehaJungsanPrice
	public FsumteenxteenrJungsanPrice
	public FsumteenteenspJungsanPrice
	public FsumteenteenapJungsanPrice
	public FsumKCTEN0001mJungsanPrice
	public FsumnaverpayJungsanPrice
	public FsumnaverpayJungsanPoint
	public FsumpaycoJungsanPrice
	public FsumbankipkumJungsanPrice
	public Fsumbankipkum_10x10JungsanPrice
	public Fsumbankipkum_fingersJungsanPrice
	public FsumbankrefundJungsanPrice
	public Fsumbankrefund_10x10JungsanPrice
	public Fsumbankrefund_fingersJungsanPrice
	public Fsum10x10_2JungsanPrice
	public FsumR5523JungsanPrice
	public FsummobiliansJungsanPrice
	public FsumPGgifticonJungsanPrice
	public FsumPGgifttingJungsanPrice
	public FsumPGokcashbagJungsanPrice
	public FsumKakaopayPrice
	public FsumKakaoJungsanPrice
	public FsumPGtossJungsanPrice
	public FsumPGchaiJungsanPrice
	public FsumPGConvinienspayJungsanPrice

    public FmeachulPrice
    public FetcPrice

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CPGDataAdvPriceItem

	public Fyyyymm
	public FtargetGbn
	public FPGgubun
	public FPGuserid
	public FappPrice
	public FtotAdvPrice
	public FpayLogAdvPrice
	public FgiftCardAdvPrice
	public FdepositAdvPrice
	public FetcPrice
	public Fregdate
	public FpayReqPrice

	public FreasonGubunALL
	public FreasonGubun001
	public FreasonGubun002
    public FreasonGubun003
    public FreasonGubun004
	public FreasonGubun020
	public FreasonGubun025
	public FreasonGubun030
	public FreasonGubun035
	public FreasonGubun040
	public FreasonGubun950
	public FreasonGubun999
	public FreasonGubun900
	public FreasonGubun901
	public FreasonGubun800
	public FreasonGubunXXX

	public function ShowDiffSUM()
		dim reasonGubunSUM : reasonGubunSUM = 0

		if Not IsNull(FreasonGubun001) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun001
		else
			reasonGubunSUM = reasonGubunSUM + FpayLogAdvPrice
		end if

		if Not IsNull(FreasonGubun002) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun002
		end if

		if Not IsNull(FreasonGubun020) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun020
		end if

		if Not IsNull(FreasonGubun025) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun025
		end if

		if Not IsNull(FreasonGubun030) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun030
		end if

		if Not IsNull(FreasonGubun035) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun035
		end if

		if Not IsNull(FreasonGubun040) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun040
		end if

		if Not IsNull(FreasonGubun950) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun950
		end if

		if Not IsNull(FreasonGubun999) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun999
		end if

		if Not IsNull(FreasonGubun900) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun900
		end if

		if Not IsNull(FreasonGubun901) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun901
		end if

		if Not IsNull(FreasonGubun800) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun800
		end if

		if Not IsNull(FreasonGubunXXX) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubunXXX
		end if

		if (GetTotalAdvPrice() - reasonGubunSUM) <> 0 then
			response.write "<font color='red'>" & FormatNumber((GetTotalAdvPrice() - reasonGubunSUM), 0) & "</font>"
		end if
	end function

	public function ShowDiffIfExist(val1, val2)
		if IsNull(val1) or IsNull(val2) then
			exit function
		end if

		if (val2 = 0) then
			exit function
		end if

		if (val1 <> val2) then
			response.write "<font color='red'><acronym title='확인사항 :<br />1. 주문번호 입금월이 지난달인지 체크'>(" & FormatNumber((val1 - val2), 0) & ")</acronym></font>"
		end if
	end function

	public function ShowDiffIfExistWithPGgubun(pggubun, val1, val2)
		if IsNull(val1) or IsNull(val2) then
			exit function
		end if

        select case pggubun
            case "balance", "giftcard", "nopayment", "mileage", "CASH":
                '// 승인내역에 내역 안올라가는 건들
                exit function
            case else:
                ''
        end select

		if (val1 <> val2) then
			response.write "<br /><font color='red'><acronym title='확인사항 :<br />1. 주문번호 입금월이 지난달인지 체크'>(" & FormatNumber((val1 - val2), 0) & ")</acronym></font>"
		end if
	end function

	public function GetMeachulPrice(pggubun, val1, val2)
		if IsNull(val1) or IsNull(val2) then
			GetMeachulPrice = 0
		end if

        select case pggubun
            case "balance", "giftcard", "nopayment", "mileage", "CASH":
                '// 승인내역에 내역 안올라가는 건들
                GetMeachulPrice = val1
            case else:
                GetMeachulPrice = val2
        end select
	end function

	public function GetDiffIfExist(val1, val2)
		GetDiffIfExist = ""
		if IsNull(val1) or IsNull(val2) then
			exit function
		end if

		if (val2 = 0) then
			exit function
		end if

		if (val1 <> val2) then
			GetDiffIfExist = "<font color='red'>(" & FormatNumber((val1 - val2), 0) & ")</font>"
		end if
	end function

	public function GetTotalAdvPrice
		GetTotalAdvPrice = FpayLogAdvPrice + FgiftCardAdvPrice + FdepositAdvPrice + FetcPrice
	end function

	public function GetAdvPriceSUM
		dim reasonGubunSUM : reasonGubunSUM = 0

		if Not IsNull(FreasonGubunALL) and (FreasonGubunALL <> 0) then
			GetAdvPriceSUM = FreasonGubunALL
			exit function
		end if

		if Not IsNull(FreasonGubun001) and (FreasonGubun001 <> 0) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun001
		else
			reasonGubunSUM = reasonGubunSUM + FpayLogAdvPrice
		end if

		if Not IsNull(FreasonGubun002) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun002
		end if

		if Not IsNull(FreasonGubun020) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun020
		end if

		if Not IsNull(FreasonGubun025) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun025
		end if

		if Not IsNull(FreasonGubun030) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun030
		end if

		if Not IsNull(FreasonGubun035) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun035
		end if

		if Not IsNull(FreasonGubun040) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun040
		end if

		if Not IsNull(FreasonGubun950) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun950
		end if

		if Not IsNull(FreasonGubun999) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun999
		end if

		if Not IsNull(FreasonGubun900) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun900
		end if

		if Not IsNull(FreasonGubun901) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun901
		end if

		if Not IsNull(FreasonGubun800) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubun800
		end if

		if Not IsNull(FreasonGubunXXX) then
			reasonGubunSUM = reasonGubunSUM + FreasonGubunXXX
		end if

		GetAdvPriceSUM = reasonGubunSUM
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CPGData
    public FItemList()
	public FOneItem
	public FArrCardComp

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
    public FTotalAppPrice

	public FRectExcMatchFinish
	public FRectExcChargeInput
	public FRectOnlyCardPriceNotSame
	public FRectOnlyPriceNotEqual
	public FRectIncJumunLog

	public FRectDateType
	public FRectStartDate
	public FRectEndDate

	public FRectStartIpkumdate
	public FRectEndIpkumDate
	public FRectYYYYMM

	public FRectShopid
	public FRectAppDivCode
	public FRectCardReaderID
	public FRectCardGubun
	public FRectCardComp
	public FRectCardAffiliateNo
	public FRectIpkumdate

	public FRectSearchField
	public FRectSearchText

	public FRectIdx

	public FRectSiteName
	public FRectPGuserid
	public FRectAppMethod
	public FRectDateGubun
	public FRectPGGubun

	public FRectReasonGubun

	public FRectShowJumunLog
	public FRectShowJumunLogNotMatch

	public function getPGDataAdvPriceList()
	    dim i,sqlStr, addSqlStr, startDate, endDate

		addSqlStr = ""

	    ''if (FRectYYYYMM <> "") then
    	''    addSqlStr = addSqlStr + " and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
    	''end if

		startDate = FRectYYYYMM + "-01"
		endDate = Left(DateAdd("m", 1, startDate), 10)

		if (endDate > LEFT(Now(), 10)) then
			endDate = LEFT(Now(), 10)
		end if
		'response.write sqlstr & "<Br>"

		sqlStr = " select top 100 " & vbCrLf
		sqlStr = sqlStr + " 	s.* " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubunALL, 0) as reasonGubunALL " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun001, 0) as reasonGubun001 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun002, 0) as reasonGubun002 " & vbCrLf
        sqlStr = sqlStr + " 	, IsNull(T.reasonGubun003, 0) as reasonGubun003 " & vbCrLf
        sqlStr = sqlStr + " 	, IsNull(T.reasonGubun004, 0) as reasonGubun004 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun020, 0) as reasonGubun020 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun025, 0) as reasonGubun025 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun030, 0) + (case when s.targetGbn = 'OF' and s.yyyymm >= '2018-01' then giftCardAdvPrice else 0 end) as reasonGubun030 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun035, 0) as reasonGubun035 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun040, 0) as reasonGubun040 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun900, 0) + (case when s.targetGbn = 'OF' and s.yyyymm >= '2015-03' then etcPrice else 0 end) as reasonGubun900 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun901, 0) as reasonGubun901 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun800, 0) as reasonGubun800 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun950, 0) as reasonGubun950 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubun999, 0) as reasonGubun999 " & vbCrLf
		sqlStr = sqlStr + " 	, IsNull(T.reasonGubunXXX, 0) as reasonGubunXXX " & vbCrLf
		sqlStr = sqlStr + " from db_summary.dbo.tbl_appPrc_advPrc_Sum s with (nolock)" & vbCrLf
		sqlStr = sqlStr + " 	left join ( " & vbCrLf

		sqlStr = sqlStr + " 		select " & vbCrLf
		sqlStr = sqlStr + " 			(case " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('bankipkum_10x10', 'bankrefund_10x10', 'tenbyten01', 'tenbyten02') or pggubun in ('gifticon', 'giftting', 'inicis', 'newkakaopay', 'naverpay', 'payco', 'toss', 'chai', 'inirental', 'okcashbag', 'convinienspay') and (pguserid <> 'teenxteen3') then 'ON' " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('') or pggubun in ('uplus') then 'OF' " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('bankrefund_fingers', 'bankipkum_fingers', 'teenxteen3') or pggubun in ('kcp') then 'AC' " & vbCrLf
		sqlStr = sqlStr + " 				else 'XX' end) as targetGbn, " & vbCrLf
		sqlStr = sqlStr + " 			pggubun, " & vbCrLf
		sqlStr = sqlStr + " 			pguserid, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(appprice), 0) as reasonGubunALL, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '001' then appprice else 0 end), 0) as reasonGubun001, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '002' then appprice else 0 end), 0) as reasonGubun002, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '003' then appprice else 0 end), 0) as reasonGubun003, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '004' then appprice else 0 end), 0) as reasonGubun004, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '020' then appprice else 0 end), 0) as reasonGubun020, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '025' then appprice else 0 end), 0) as reasonGubun025, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '030' then appprice else 0 end), 0) as reasonGubun030, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '035' then appprice else 0 end), 0) as reasonGubun035, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '040' then appprice else 0 end), 0) as reasonGubun040, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '900' then appprice else 0 end), 0) as reasonGubun900, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '901' then appprice else 0 end), 0) as reasonGubun901, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '800' then appprice else 0 end), 0) as reasonGubun800, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '950' then appprice else 0 end), 0) as reasonGubun950, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '999' then appprice else 0 end), 0) as reasonGubun999, " & vbCrLf
		sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') then appprice else 0 end), 0) as reasonGubunXXX " & vbCrLf
		sqlStr = sqlStr + " 		from db_order.dbo.tbl_onlineApp_log with (nolock)" & vbCrLf
		sqlStr = sqlStr + " 		where " & vbCrLf
		sqlStr = sqlStr + " 			1 = 1 " & vbCrLf
		sqlStr = sqlStr + " 			and IsNull(canceldate, appdate) >= '" + CStr(startDate) + "' " & vbCrLf
		sqlStr = sqlStr + " 			and IsNull(canceldate, appdate) < '" + CStr(endDate) + "' " & vbCrLf
		sqlStr = sqlStr + " 		group by " & vbCrLf
		sqlStr = sqlStr + " 			(case " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('bankipkum_10x10', 'bankrefund_10x10', 'tenbyten02') or pggubun in ('gifticon', 'giftting', 'inicis', 'newkakaopay', 'naverpay', 'payco', 'toss', 'chai', 'inirental', 'okcashbag', 'convinienspay') and (pguserid <> 'teenxteen3') then 'ON' " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('') or pggubun in ('uplus') then 'OF' " & vbCrLf
		sqlStr = sqlStr + " 				when pguserid in ('bankrefund_fingers', 'bankipkum_fingers', 'teenxteen3') or pggubun in ('kcp') then 'AC' " & vbCrLf
		sqlStr = sqlStr + " 				else 'XX' end), " & vbCrLf
		sqlStr = sqlStr + " 			pggubun, pguserid " & vbCrLf

        sqlStr = sqlStr + " 			union all  " & vbCrLf

        sqlStr = sqlStr + " 		select " & vbCrLf
        sqlStr = sqlStr + " 			targetGbn, " & vbCrLf
        sqlStr = sqlStr + " 			pggubun, " & vbCrLf
        sqlStr = sqlStr + " 			pguserid, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(appprice), 0) as reasonGubunALL, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '001' then appprice else 0 end), 0) as reasonGubun001, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '002' then appprice else 0 end), 0) as reasonGubun002, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '003' then appprice else 0 end), 0) as reasonGubun003, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '004' then appprice else 0 end), 0) as reasonGubun004, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '020' then appprice else 0 end), 0) as reasonGubun020, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '025' then appprice else 0 end), 0) as reasonGubun025, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '030' then appprice else 0 end), 0) as reasonGubun030, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '035' then appprice else 0 end), 0) as reasonGubun035, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '040' then appprice else 0 end), 0) as reasonGubun040, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '900' then appprice else 0 end), 0) as reasonGubun900, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '901' then appprice else 0 end), 0) as reasonGubun901, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '800' then appprice else 0 end), 0) as reasonGubun800, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '950' then appprice else 0 end), 0) as reasonGubun950, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') = '999' then appprice else 0 end), 0) as reasonGubun999, " & vbCrLf
        sqlStr = sqlStr + " 			IsNull(sum(case when IsNull(reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') then appprice else 0 end), 0) as reasonGubunXXX " & vbCrLf
        sqlStr = sqlStr + " 		from db_summary.dbo.tbl_appPrc_advPrc_Sum_OFF with (nolock)" & vbCrLf
        sqlStr = sqlStr + " 		where " & vbCrLf
        sqlStr = sqlStr + " 			1 = 1 " & vbCrLf
        sqlStr = sqlStr + " 			and yyyymm = '" & FRectYYYYMM & "' " & vbCrLf
        sqlStr = sqlStr + " 		group by " & vbCrLf
        sqlStr = sqlStr + " 			targetGbn, pggubun, pguserid " & vbCrLf

		sqlStr = sqlStr + " 	) T " & vbCrLf
		sqlStr = sqlStr + " 	on " & vbCrLf
		sqlStr = sqlStr + " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr + " 		and s.targetGbn = T.targetGbn " & vbCrLf
		sqlStr = sqlStr + " 		and s.pggubun = T.pggubun " & vbCrLf
		sqlStr = sqlStr + " 		and s.pguserid = T.pguserid " & vbCrLf
		sqlStr = sqlStr + " where 1 = 1 " & vbCrLf
		sqlStr = sqlStr + " and s.yyyymm = '" + CStr(FRectYYYYMM) + "' " & vbCrLf
		sqlStr = sqlStr + " order by s.targetGbn desc, s.pggubun, s.pguserid " & vbCrLf

		''response.write sqlStr & "<Br>"
	    rsget.pagesize = 100
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		FTotalPage = 1

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = 1
			do until rsget.eof
				set FItemList(i) = new CPGDataAdvPriceItem

				FItemList(i).Fyyyymm				= rsget("yyyymm")
				FItemList(i).FtargetGbn				= rsget("targetGbn")
				FItemList(i).FPGgubun				= rsget("PGgubun")
				FItemList(i).FPGuserid				= rsget("PGuserid")
				FItemList(i).FappPrice				= rsget("appPrice")
				FItemList(i).FtotAdvPrice			= rsget("totAdvPrice")
				FItemList(i).FpayLogAdvPrice		= rsget("payLogAdvPrice")
				FItemList(i).FgiftCardAdvPrice		= rsget("giftCardAdvPrice")
				FItemList(i).FdepositAdvPrice		= rsget("depositAdvPrice")
				FItemList(i).FetcPrice				= rsget("etcPrice")
				FItemList(i).Fregdate				= rsget("regdate")
				FItemList(i).FpayReqPrice			= rsget("payReqPrice")

				FItemList(i).FreasonGubunALL		= rsget("reasonGubunALL")
				FItemList(i).FreasonGubun001		= rsget("reasonGubun001")
				FItemList(i).FreasonGubun002		= rsget("reasonGubun002")
                FItemList(i).FreasonGubun003		= rsget("reasonGubun003")
                FItemList(i).FreasonGubun004		= rsget("reasonGubun004")
				FItemList(i).FreasonGubun020		= rsget("reasonGubun020")
				FItemList(i).FreasonGubun025		= rsget("reasonGubun025")
				FItemList(i).FreasonGubun030		= rsget("reasonGubun030")
				FItemList(i).FreasonGubun035		= rsget("reasonGubun035")
				FItemList(i).FreasonGubun040		= rsget("reasonGubun040")
				FItemList(i).FreasonGubun950		= rsget("reasonGubun950")
				FItemList(i).FreasonGubun999		= rsget("reasonGubun999")
				FItemList(i).FreasonGubun900		= rsget("reasonGubun900")
				FItemList(i).FreasonGubun901		= rsget("reasonGubun901")
				FItemList(i).FreasonGubun800		= rsget("reasonGubun800")
				FItemList(i).FreasonGubunXXX		= rsget("reasonGubunXXX")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function getPGDataList_OFF()
	    dim i,sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

	    if (FRectPGGubun <> "") then
    	    addSqlStr = addSqlStr + " and m.pggubun = '" + CStr(FRectPGGubun) + "' "
    	end if

	    if (FRectExcMatchFinish <> "") then
    	    addSqlStr = addSqlStr + " and m.shopJumunMasterIdx is NULL "
    	end if

	    if (FRectExcChargeInput <> "") then
    	    addSqlStr = addSqlStr + " and IsNull(m.cardChargePrice,0) = 0 "
			'// addSqlStr = addSqlStr + " and m.cardComp <> '비씨카드사' "
    	end if

		Select Case FRectDateType
			Case "B"
				'// 입금예정일
				if FRectStartdate <> "" then
					addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartdate) + "'"
				end if
				if FRectEndDate <> "" then
					addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndDate) + "'"
				end if
			Case Else
				'// 거래일자
				if FRectStartdate <> "" then
					addSqlStr = addSqlStr + " and m.appDate>='" + CStr(FRectStartdate) + "'"
				end if
				if FRectEndDate <> "" then
					addSqlStr = addSqlStr + " and m.appDate<'" + CStr(FRectEndDate) + "'"
				end if
		End Select

		if (FRectshopid <> "") then
			addSqlStr = addSqlStr + " and m.shopid = '" + CStr(FRectshopid) + "' "
		end if

		if (FRectAppDivCode <> "") then
			addSqlStr = addSqlStr + " and m.appDivCode = '" + CStr(FRectAppDivCode) + "' "
		end if

		if (FRectCardReaderID <> "") then
			addSqlStr = addSqlStr + " and m.cardReaderID = '" + CStr(FRectCardReaderID) + "' "
		end if

		if (FRectCardGubun <> "") then
			addSqlStr = addSqlStr + " and m.cardGubun = '" + CStr(FRectCardGubun) + "' "
		end if

		if (FRectCardComp <> "") then
			addSqlStr = addSqlStr + " and m.cardComp = '" + CStr(FRectCardComp) + "' "
		end if

		if (FRectCardAffiliateNo <> "") then
			addSqlStr = addSqlStr + " and m.cardAffiliateNo = '" + CStr(FRectCardAffiliateNo) + "' "
		end if

		if (FRectIpkumdate <> "") then
			addSqlStr = addSqlStr + " and m.ipkumdate = '" + CStr(FRectIpkumdate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField <> "orderCardPrice") then
				addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			elseif (FRectSearchField = "orderCardPrice") then
				addSqlStr = addSqlStr + " and IsNull(s.cardsum, 0) = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectReasonGubun <> "") then
			if (FRectReasonGubun = "XXX") then
				addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') "
			else
				addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') = '" + CStr(FRectReasonGubun) + "' "
			end if
		end if

		if (FRectOnlyCardPriceNotSame <> "") then
			addSqlStr = addSqlStr + " and m.cardPrice <> IsNull(s.cardsum, m.cardPrice) "
		end if

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
	    sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cardApp_log m"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopjumun_master s "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.orderserial = s.orderno "
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		'response.write sqlstr & "<Br>"
    	rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		'// ====================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, s.cardsum as orderCardPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log m "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopjumun_master s "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.orderserial = s.orderno "
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by m.appDate desc"

		response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPGDataItem

				FItemList(i).Fidx					= rsget("idx")
				FItemList(i).FPGgubun				= rsget("PGgubun")
				FItemList(i).FPGkey					= rsget("PGkey")
				FItemList(i).FappDivCode			= rsget("appDivCode")
				FItemList(i).FappDate				= rsget("appDate")
				FItemList(i).FcardReaderID			= rsget("cardReaderID")
				FItemList(i).FcardPrice				= rsget("cardPrice")
				FItemList(i).FcardAppNo				= rsget("cardAppNo")
				FItemList(i).Fshopid				= rsget("shopid")
				FItemList(i).FshopJumunMasterIdx	= rsget("shopJumunMasterIdx")
				FItemList(i).Fregdate				= rsget("regdate")

				FItemList(i).FcardGubun				= rsget("cardGubun")
				FItemList(i).FcardComp				= rsget("cardComp")
				FItemList(i).FcardAffiliateNo		= rsget("cardAffiliateNo")

				FItemList(i).Fipkumdate				= rsget("ipkumdate")
				FItemList(i).FipkumPrice			= rsget("ipkumPrice")
				FItemList(i).FcardChargePrice		= rsget("cardChargePrice")

				FItemList(i).FreasonGubun			= rsget("reasonGubun")

				FItemList(i).Forderserial			= rsget("orderserial")
				FItemList(i).ForderCardPrice		= rsget("orderCardPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function getPGDataOne_OFF()
	    dim i,sqlStr

		'// ====================================================================
		sqlStr = "select top 1 m.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log m "
	    sqlStr = sqlStr + " where m.idx = " + CStr(FRectIdx)

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneItem = new CPGDataItem

		if  not rsget.EOF  then
			FOneItem.Fidx					= rsget("idx")
			FOneItem.FPGgubun				= rsget("PGgubun")
			FOneItem.FPGkey					= rsget("PGkey")
			FOneItem.FappDivCode			= rsget("appDivCode")
			FOneItem.FappDate				= rsget("appDate")
			FOneItem.FcardReaderID			= rsget("cardReaderID")
			FOneItem.FcardPrice				= rsget("cardPrice")
			FOneItem.FcardAppNo				= rsget("cardAppNo")
			FOneItem.Fshopid				= rsget("shopid")
			FOneItem.FshopJumunMasterIdx	= rsget("shopJumunMasterIdx")
			FOneItem.Fregdate				= rsget("regdate")

			FOneItem.FcardGubun				= rsget("cardGubun")
			FOneItem.FcardComp				= rsget("cardComp")
			FOneItem.FcardAffiliateNo		= rsget("cardAffiliateNo")
			FOneItem.Fipkumdate				= rsget("ipkumdate")
			FOneItem.Forderserial			= rsget("orderserial")
		end if
		rsget.Close
    end function

	public function getPGDataStatisticList_OFF()
		dim i, j, sqlStr, addSqlStr
		dim tmpArrCardComp
		dim tmpStr, tmpVal

		'' sqlStr = " select distinct cardComp "
		'' sqlStr = sqlStr + " from "
		'' sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
		'' sqlStr = sqlStr + " where "
		'' sqlStr = sqlStr + " 	1 = 1 "

		'' if (FRectShopid <> "") then
		'' 	sqlStr = sqlStr + " 	and l.shopid = '" + CStr(FRectShopid) + "' "
		'' end if

		'' if (FRectDateGubun = "ipkumdate") then
		'' 	sqlStr = sqlStr + " 	and l.ipkumdate >= '" + CStr(FRectStartdate) + "' "
		'' 	sqlStr = sqlStr + " 	and l.ipkumdate < '" + CStr(FRectEndDate) + "' "
		'' else
		'' 	sqlStr = sqlStr + " 	and l.appDate >= '" + CStr(FRectStartdate) + "' "
		'' 	sqlStr = sqlStr + " 	and l.appDate < '" + CStr(FRectEndDate) + "' "
		'' end if

		'' sqlStr = sqlStr + " order by "
		'' sqlStr = sqlStr + " 	cardComp "
		'' rsget.Open sqlStr,dbget,1

		'' FArrCardComp = ""
		'' if  not rsget.EOF  then
		'' 	do until rsget.eof
		'' 		if (FArrCardComp = "") then
		'' 			FArrCardComp = rsget("cardComp")
		'' 		else
		'' 			FArrCardComp = FArrCardComp + "|" + rsget("cardComp")
		'' 		end if

		'' 		rsget.moveNext
		'' 	loop
		'' end if
		'' rsget.Close

		'// 카드사 고정
		'// TODO : [db_shop].[dbo].[usp_Ten_getPGDataStatisticList_OFF] 도 같이 수정해야 한다.
		FArrCardComp = "KB국민카드|NH농협카드|롯데카드사|비씨카드사|삼성카드사|신한카드|하나카드|하나SK카드|외환카드사|현대카드사|Alipay|기타"

		'// ====================================================================
		tmpArrCardComp = GetArrCardComp()

		'' sqlStr = " select T1.* "
		'' if (FRectDateGubun = "ipkumdate") then
		'' 	sqlStr = sqlStr + " 	, 0 as scmTotCardPrice "
		'' else
		'' 	sqlStr = sqlStr + " 	, isNULL(T2.scmTotCardPrice,0) as scmTotCardPrice "
		'' end if
		'' sqlStr = sqlStr + " from "
		'' sqlStr = sqlStr + " 	( "
		'' sqlStr = sqlStr + " 		select "

		'' if (FRectDateGubun = "ipkumdate") then
		'' 	sqlStr = sqlStr + " 		l.ipkumdate AS yyyymmdd "
		'' else
		'' 	sqlStr = sqlStr + " 		convert(VARCHAR(10), l.appDate, 127) AS yyyymmdd "
		'' end if

		'' for i = 0 to UBound(tmpArrCardComp)
		'' 	sqlStr = sqlStr + " 	, sum(case when cardComp = '" + CStr(tmpArrCardComp(i)) + "' then IsNull(cardPrice, 0) else 0 end) as '" + CStr(tmpArrCardComp(i)) + "' "
		'' 	sqlStr = sqlStr + " 	, sum(case when cardComp = '" + CStr(tmpArrCardComp(i)) + "' then IsNull(ipkumPrice, 0) else 0 end) as '" + CStr(tmpArrCardComp(i)) + "IPKUM' "
		'' next
		'' sqlStr = sqlStr + " 			, sum(IsNull(cardPrice, 0)) AS totCardPrice "
		'' sqlStr = sqlStr + " 			, sum(IsNull(ipkumPrice, 0)) AS totCardIpkumPrice "

		'' if (FRectShopid <> "") then
		'' 	sqlStr = sqlStr + " 			, l.shopid "
		'' end if

		'' sqlStr = sqlStr + " 		from db_shop.dbo.tbl_shopjumun_cardApp_log l "
		'' sqlStr = sqlStr + " 		where 1 = 1 "

		'' if (FRectShopid <> "") then
		'' 	sqlStr = sqlStr + " 			AND l.shopid = '" + CStr(FRectShopid) + "' "
		'' end if

		'' if (FRectDateGubun = "ipkumdate") then
		'' 	sqlStr = sqlStr + " 	and l.ipkumdate >= '" + CStr(FRectStartdate) + "' "
		'' 	sqlStr = sqlStr + " 	and l.ipkumdate < '" + CStr(FRectEndDate) + "' "
		'' else
		'' 	sqlStr = sqlStr + " 	and l.appDate >= '" + CStr(FRectStartdate) + "' "
		'' 	sqlStr = sqlStr + " 	and l.appDate < '" + CStr(FRectEndDate) + "' "
		'' end if

		'' sqlStr = sqlStr + " 		group by "

		'' if (FRectDateGubun = "ipkumdate") then
		'' 	sqlStr = sqlStr + " 		l.ipkumdate "
		'' else
		'' 	sqlStr = sqlStr + " 		convert(VARCHAR(10), l.appDate, 127) "
		'' end if

		'' if (FRectShopid <> "") then
		'' 	sqlStr = sqlStr + " 			, l.shopid "
		'' end if

		'' sqlStr = sqlStr + " 	) T1 "

		'' if (FRectDateGubun = "ipkumdate") then
		'' 	''
		'' else
		'' 	sqlStr = sqlStr + " 	left join ( "
		'' 	sqlStr = sqlStr + " 		select convert(VARCHAR(10), m.shopregdate, 121) AS yyyymmdd "
		'' 	sqlStr = sqlStr + " 			,sum(cardsum) AS scmTotCardPrice "

		'' 	if (FRectShopid <> "") then
		'' 		sqlStr = sqlStr + " 			,m.shopid "
		'' 	end if

		'' 	sqlStr = sqlStr + " 		from db_shop.dbo.tbl_shopjumun_master m "
		'' 	sqlStr = sqlStr + " 		where 1 = 1 "
		'' 	sqlStr = sqlStr + " 			AND m.cancelyn = 'N' "

		'' 	if (FRectShopid <> "") then
		'' 		sqlStr = sqlStr + " 			AND m.shopid = '" + CStr(FRectShopid) + "' "
		'' 	else
		'' 		sqlStr = sqlStr + " 		AND m.shopid in ( "
		'' 		sqlStr = sqlStr + " 			select distinct l.shopid "
		'' 		sqlStr = sqlStr + " 			from db_shop.dbo.tbl_shopjumun_cardApp_log l "
		'' 		sqlStr = sqlStr + " 			where l.appDate>='" + CStr(FRectStartdate) + "' and l.appDate< '" + CStr(FRectEndDate) + "' and l.shopid is not NULL "
		'' 		sqlStr = sqlStr + " 		) "
		'' 	end if

		'' 	sqlStr = sqlStr + " 			AND m.shopregdate >= '" + CStr(FRectStartdate) + "' "
		'' 	sqlStr = sqlStr + " 			AND m.shopregdate < '" + CStr(FRectEndDate) + "' "
		'' 	sqlStr = sqlStr + " 		group by convert(VARCHAR(10), m.shopregdate, 121) "

		'' 	if (FRectShopid <> "") then
		'' 		sqlStr = sqlStr + " 			,m.shopid "
		'' 	end if

		'' 	sqlStr = sqlStr + " 	) T2 "
		'' 	sqlStr = sqlStr + " 	on "
		'' 	sqlStr = sqlStr + " 		1 = 1 "

		'' 	if (FRectShopid <> "") then
		'' 		sqlStr = sqlStr + " 		and T1.shopid = T2.shopid "
		'' 	end if

		'' 	sqlStr = sqlStr + " 		and T1.yyyymmdd = T2.yyyymmdd "
		'' end if

		'' sqlStr = sqlStr + " order by "
		'' sqlStr = sqlStr + " 	T1.yyyymmdd "

		'' if (FRectShopid <> "") then
		'' 	sqlStr = sqlStr + " 	, T1.shopid "
		'' end if


		sqlStr = " exec [db_shop].[dbo].[usp_Ten_getPGDataStatisticList_OFF] '" + CStr(FRectShopid) + "', '" + CStr(FRectDateGubun) + "', '" + CStr(FRectStartdate) + "', '" + CStr(FRectEndDate) + "', '" + CStr(FRectStartIpkumdate) + "', '" + CStr(FRectEndIpkumDate) + "', '" + CStr(FRectReasonGubun) + "', '" + CStr(FRectPGGubun) + "', '" + CStr(FRectPGuserid) + "' "
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1


		''response.write sqlStr & "<Br>"
		''rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalcount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CPGDataStatisticItem

				if (FRectShopid <> "") then
					FItemList(i).Fshopid = rsget("shopid")
				end if

				FItemList(i).Fyyyymmdd = rsget("yyyymmdd")


				'// ============================================================
				tmpStr = ""
				for j = 0 to UBound(tmpArrCardComp)
					tmpVal = rsget(tmpArrCardComp(j))
					if IsNull(tmpVal) then
						tmpVal = 0
					end if
					if (tmpStr = "") then
						tmpStr = CStr(tmpVal)
					else
						tmpStr = tmpStr + "|" + CStr(tmpVal)
					end if
				next
				FItemList(i).FarrSumCardPrice		= Split(tmpStr, "|")

				'// ============================================================
				tmpStr = ""
				for j = 0 to UBound(tmpArrCardComp)
					tmpVal = rsget(tmpArrCardComp(j) + "IPKUM")
					if IsNull(tmpVal) then
						tmpVal = 0
					end if
					if (tmpStr = "") then
						tmpStr = CStr(tmpVal)
					else
						tmpStr = tmpStr + "|" + CStr(tmpVal)
					end if
				next
				FItemList(i).FarrSumCardIpkumPrice	= Split(tmpStr, "|")


				FItemList(i).FtotSumCardPrice		= rsget("totCardPrice")
				FItemList(i).FtotSumCardIpkumPrice	= rsget("totCardIpkumPrice")
				FItemList(i).FscmTotCardPrice		= rsget("scmTotCardPrice")
				FItemList(i).FcardPriceNotMatch		= rsget("CardPriceNotMatch")

                FItemList(i).FmeachulPrice			= rsget("meachulPrice")
                FItemList(i).FetcPrice				= rsget("etcPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function getPGDataList_ON()
	    dim i,sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

	    if (FRectPGGubun <> "") then
    	    addSqlStr = addSqlStr + " and m.pggubun = '" + CStr(FRectPGGubun) + "' "
    	end if

	    if (FRectExcMatchFinish <> "") then
			addSqlStr = addSqlStr + " and ( "
			addSqlStr = addSqlStr + " 	(m.appDivCode = 'A' and IsNull(m.orderserial, '') = '') "
			addSqlStr = addSqlStr + " 	or "
			addSqlStr = addSqlStr + " 	(m.appDivCode <> 'A' and (m.csasid is NULL or IsNull(m.orderserial, '') = '')) "
			addSqlStr = addSqlStr + " ) "
			''addSqlStr = addSqlStr + " and NOT (m.pggubun='naverpay' and LEN(m.pgkey)>20) "
			''addSqlStr = addSqlStr + " and not (m.appDivCode = 'C' and m.pgcskey = 'CANCELALL' and m.orderserial is not NULL) "
    	end if

		'// 승인일자
		if FRectStartdate <> "" then
			addSqlStr = addSqlStr + " and IsNull(m.cancelDate, m.appDate)>='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate <> "" then
			addSqlStr = addSqlStr + " and IsNull(m.cancelDate, m.appDate)<'" + CStr(FRectEndDate) + "'"
		end if

		'// 입금예정일
		if FRectStartIpkumdate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartIpkumdate) + "'"
		end if
		if FRectEndIpkumDate <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndIpkumDate) + "'"
		end if

		if (FRectSiteName <> "") then
			addSqlStr = addSqlStr + " and m.sitename = '" + CStr(FRectSiteName) + "' "
		end if

		if (FRectAppDivCode <> "") then
			addSqlStr = addSqlStr + " and m.appDivCode = '" + CStr(FRectAppDivCode) + "' "
		end if

		if (FRectIpkumdate <> "") then
			addSqlStr = addSqlStr + " and m.ipkumdate = '" + CStr(FRectIpkumdate) + "' "
		end if

		if (FRectPGuserid <> "") then
			addSqlStr = addSqlStr + " and m.PGuserid = '" + CStr(FRectPGuserid) + "' "
		end if

		if (FRectAppMethod <> "") then
			addSqlStr = addSqlStr + " and m.appMethod = '" + CStr(FRectAppMethod) + "' "
		end if

		if (FRectReasonGubun <> "") then
			if (FRectReasonGubun = "XXX") then
				addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') "
			else
				addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') = '" + CStr(FRectReasonGubun) + "' "
			end if
		end if

		if (FRectOnlyPriceNotEqual <> "") then
			addSqlStr = addSqlStr + " and m.appdivcode = 'A' "
			addSqlStr = addSqlStr + " and e.acctamount <> m.appprice "
		end if

		if (FRectShowJumunLogNotMatch = "Y") then
			FRectShowJumunLog = "Y"
			addSqlStr = addSqlStr + " and (p.pggubun is NULL or (p.pggubun is not NULL and DateDiff(month, IsNull(m.cancelDate, m.appDate), p.payDate) <> 0)) "
			''addSqlStr = addSqlStr + " and m.sitename not in ('fingers', '10x10gift') "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField = "orderserial") then
				addSqlStr = addSqlStr + " and Left(m.orderserial, 11) = '" + CStr(Left(FRectSearchText, 11)) + "' "
			else
				addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg, IsNull(sum(appPrice), 0) as totAppPrice "
	    sqlStr = sqlStr + " from db_order.dbo.tbl_onlineApp_log m"

		if (FRectOnlyPriceNotEqual <> "") then
			sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.orderserial = e.orderserial "
			sqlStr = sqlStr + " 	and e.acctdiv = m.appmethod "
		end if

		if (FRectShowJumunLog = "Y") then
			sqlStr = sqlStr + " left join db_datamart.dbo.tbl_order_payment_log p "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.pggubun = p.pggubun "
			sqlStr = sqlStr + " 	and m.pgkey = p.pgkey "
			sqlStr = sqlStr + " 	and m.pgcskey = p.pgcskey "
			sqlStr = sqlStr + " 	and m.appprice = p.realPayPrice "
		end if

	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		''response.write sqlstr & "<Br>"

		if (FRectOnlyPriceNotEqual <> "") or (FRectShowJumunLog = "Y") then
			'// 77번 디비
			db3_rsget.Open sqlStr,db3_dbget,1
				FTotalCount = db3_rsget("cnt")
				FTotalPage = db3_rsget("totPg")
                FTotalAppPrice = db3_rsget("totAppPrice")
			db3_rsget.Close
		else
			rsget.Open sqlStr,dbget,1
				FTotalCount = rsget("cnt")
				FTotalPage = rsget("totPg")
                FTotalAppPrice = rsget("totAppPrice")
			rsget.Close
		end if

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			''exit function
		end if

		'// ====================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
		if (FRectShowJumunLog = "Y") then
			sqlStr = sqlStr + " , p.orderserial as logorderserial, p.suborderserial as logsuborderserial "
		end if
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log m "

		if (FRectOnlyPriceNotEqual <> "") then
			sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.orderserial = e.orderserial "
			sqlStr = sqlStr + " 	and e.acctdiv = m.appmethod "
		end if

		if (FRectShowJumunLog = "Y") then
			sqlStr = sqlStr + " left join db_datamart.dbo.tbl_order_payment_log p "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.pggubun = p.pggubun "
			sqlStr = sqlStr + " 	and m.pgkey = p.pgkey "
			sqlStr = sqlStr + " 	and m.pgcskey = p.pgcskey "
			sqlStr = sqlStr + " 	and m.appprice = p.realPayPrice "
		end if

	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by IsNull(m.cancelDate, m.appDate) desc, m.pgkey desc, m.pgcskey desc"

		if session("ssBctId")="tozzinet" then
		response.write sqlStr & "<Br>"
		else
		'response.write sqlStr & "<Br>"
		end if

		if (FRectOnlyPriceNotEqual <> "") or (FRectShowJumunLog = "Y") then
			db3_rsget.pagesize = FPageSize
			db3_rsget.Open sqlStr,db3_dbget,1

			FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

			redim preserve FItemList(FResultCount)
			i=0
			if  not db3_rsget.EOF  then
				db3_rsget.absolutepage = FCurrPage
				do until db3_rsget.eof
					set FItemList(i) = new CPGDataItem

					FItemList(i).Fidx					= db3_rsget("idx")
					FItemList(i).FPGgubun				= db3_rsget("PGgubun")
					FItemList(i).FPGkey					= db3_rsget("PGkey")
					FItemList(i).FPGCSkey				= db3_rsget("PGCSkey")
					FItemList(i).FappDivCode			= db3_rsget("appDivCode")
					FItemList(i).FappMethod				= db3_rsget("appMethod")
					FItemList(i).FappDate				= db3_rsget("appDate")
					FItemList(i).FcancelDate			= db3_rsget("cancelDate")
					FItemList(i).Fsitename				= db3_rsget("sitename")
					FItemList(i).Fregdate				= db3_rsget("regdate")
					FItemList(i).Fpgmeachuldate			= db3_rsget("pgmeachuldate")		'// 카드사매입일
					FItemList(i).Fipkumdate				= db3_rsget("ipkumdate")
					FItemList(i).Forderserial			= db3_rsget("orderserial")

					FItemList(i).FappPrice				= db3_rsget("appPrice")
					FItemList(i).FcommPrice				= db3_rsget("commPrice") * -1
					FItemList(i).FcommVatPrice			= db3_rsget("commVatPrice") * -1
					FItemList(i).FjungsanPrice			= db3_rsget("jungsanPrice")
					FItemList(i).Fcsasid				= db3_rsget("csasid")
					FItemList(i).FPGuserid				= db3_rsget("PGuserid")

					FItemList(i).FreasonGubun			= db3_rsget("reasonGubun")

					if (FRectShowJumunLog = "Y") then
						FItemList(i).Flogorderserial				= db3_rsget("logorderserial")
						FItemList(i).Flogsuborderserial				= db3_rsget("logsuborderserial")
					end if

					i=i+1
					db3_rsget.moveNext
				loop
			end if
			db3_rsget.Close
		else
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1

			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CPGDataItem

					FItemList(i).Fidx					= rsget("idx")
					FItemList(i).FPGgubun				= rsget("PGgubun")
					FItemList(i).FPGkey					= rsget("PGkey")
					FItemList(i).FPGCSkey				= rsget("PGCSkey")
					FItemList(i).FappDivCode			= rsget("appDivCode")
					FItemList(i).FappMethod				= rsget("appMethod")
					FItemList(i).FappDate				= rsget("appDate")
					FItemList(i).FcancelDate			= rsget("cancelDate")
					FItemList(i).Fsitename				= rsget("sitename")
					FItemList(i).Fregdate				= rsget("regdate")
					FItemList(i).Fpgmeachuldate			= rsget("pgmeachuldate")		'// 카드사매입일
					FItemList(i).Fipkumdate				= rsget("ipkumdate")
					FItemList(i).Forderserial			= rsget("orderserial")

					FItemList(i).FappPrice				= rsget("appPrice")
					FItemList(i).FcommPrice				= rsget("commPrice") * -1
					FItemList(i).FcommVatPrice			= rsget("commVatPrice") * -1
					FItemList(i).FjungsanPrice			= rsget("jungsanPrice")
					FItemList(i).Fcsasid				= rsget("csasid")
					FItemList(i).FPGuserid				= rsget("PGuserid")

					FItemList(i).FreasonGubun			= rsget("reasonGubun")

					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.Close
		end if
    end function

    public function getPGDataStatisticList_ON()
		dim i, j, sqlStr, addSqlStr

		'// ====================================================================
		sqlStr = " ;WITH T_LIST AS ( select "

		if (FRectDateGubun = "ipkumdate") then
			sqlStr = sqlStr + " 	l.ipkumdate AS yyyymmdd "
		else
			sqlStr = sqlStr + " 	convert(VARCHAR(10), IsNull(l.cancelDate, l.appDate), 127) AS yyyymmdd "
		end if
        sqlStr = sqlStr + " ,l.PGuserid,l.appMethod, sum(l.appPrice) as appPrice, sum(jungsanPrice) as jungsanPrice"
        sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l with (nolock)"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		if (FRectStartIpkumdate <> "") then
			sqlStr = sqlStr + " 	and l.ipkumdate >= '" + CStr(FRectStartIpkumdate) + "' "
			sqlStr = sqlStr + " 	and l.ipkumdate < '" + CStr(FRectEndIpkumDate) + "' "
		end if

		if (FRectStartdate <> "") then
			sqlStr = sqlStr + " 	and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' "
			sqlStr = sqlStr + " 	and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectPGuserid <> "") then
			sqlStr = sqlStr + " 	and l.PGuserid = '" + CStr(FRectPGuserid) + "' "
		end if

		if (FRectSiteName <> "") then
			if (FRectSiteName = "10x10all") then
				sqlStr = sqlStr + " and l.sitename in ('10x10', '10x10mobile') "
			else
				sqlStr = sqlStr + " and l.sitename = '" + CStr(FRectSiteName) + "' "
			end if
		end if

	    if (FRectPGGubun <> "") then
    	    sqlStr = sqlStr + " and l.pggubun = '" + CStr(FRectPGGubun) + "' "
    	end if

		if (FRectReasonGubun <> "") then
			if (FRectReasonGubun = "XXX") then
				sqlStr = sqlStr + " and IsNull(l.reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') "
			else
				sqlStr = sqlStr + " and IsNull(l.reasonGubun, '') = '" + CStr(FRectReasonGubun) + "' "
			end if
		end if

		if (FRectDateGubun = "ipkumdate") then
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	l.ipkumdate ,l.PGuserid,l.appMethod"

		else
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	convert(VARCHAR(10), IsNull(l.cancelDate, l.appDate), 127) ,l.PGuserid,l.appMethod"

		end if
		sqlStr = sqlStr + " ) " &VBCRLF

		sqlStr = sqlStr + " select yyyymmdd"
		sqlStr = sqlStr + " 	, sum(l.appPrice) as totSumPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 100 then l.appPrice else 0 end) as sumCardPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 20 then l.appPrice else 0 end) as sumBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 7 then l.appPrice else 0 end) as sumVBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 77 then l.appPrice else 0 end) as sumTenOutBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 6 then l.appPrice else 0 end) as sumTenInBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 400 then l.appPrice else 0 end) as sumHPPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 550 then l.appPrice else 0 end) as sumGifttingPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 560 then l.appPrice else 0 end) as sumGifticonPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 110 then l.appPrice else 0 end) as sumOKPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 80 then l.appPrice else 0 end) as sumAllAtPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 150 then l.appPrice else 0 end) as sumRentalPrice "

		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen3' then l.appPrice else 0 end) as sumteenxteen3Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen4' then l.appPrice else 0 end) as sumteenxteen4Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen5' then l.appPrice else 0 end) as sumteenxteen5Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen6' then l.appPrice else 0 end) as sumteenxteen6Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen8' then l.appPrice else 0 end) as sumteenxteen8Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen9' then l.appPrice else 0 end) as sumteenxteen9Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteen10' then l.appPrice else 0 end) as sumteenteen10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten01' then l.appPrice else 0 end) as sumtenbyten01Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten02' then l.appPrice else 0 end) as sumtenbyten02Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteeha' then l.appPrice else 0 end) as sumteenxteehaPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteenr' then l.appPrice else 0 end) as sumteenxteenrPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteensp' then l.appPrice else 0 end) as sumteenteenspPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteenap' then l.appPrice else 0 end) as sumteenteenapPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'KCTEN0001m' then l.appPrice else 0 end) as sumKCTEN0001mPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'newkakaopay' then l.appPrice else 0 end) as sumKakaopayPrice "

		sqlStr = sqlStr + " 	, sum(case when (l.PGuserid = 'naverpay' and l.appMethod<>80) then l.appPrice else 0 end) as sumnaverpayPrice "
		sqlStr = sqlStr + " 	, sum(case when (l.PGuserid = 'naverpay' and l.appMethod=80) then l.appPrice else 0 end) as sumnaverpayPoint "

		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'payco' then l.appPrice else 0 end) as sumpaycoPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum' then l.appPrice else 0 end) as sumbankipkumPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_10x10' then l.appPrice else 0 end) as sumbankipkum_10x10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_fingers' then l.appPrice else 0 end) as sumbankipkum_fingersPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund' then l.appPrice else 0 end) as sumbankrefundPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_10x10' then l.appPrice else 0 end) as sumbankrefund_10x10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_fingers' then l.appPrice else 0 end) as sumbankrefund_fingersPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = '10x10_2' then l.appPrice else 0 end) as sum10x10_2Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'R5523' then l.appPrice else 0 end) as sumR5523Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'mobilians' then l.appPrice else 0 end) as summobiliansPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'gifticon' then l.appPrice else 0 end) as sumPGgifticonPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'giftting' then l.appPrice else 0 end) as sumPGgifttingPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'okcashbag' then l.appPrice else 0 end) as sumPGokcashbagPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'toss' then l.appPrice else 0 end) as sumPGtossPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'chai' then l.appPrice else 0 end) as sumPGchaiPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'convinienspay' then l.appPrice else 0 end) as sumPGConvinienspayPrice "

		sqlStr = sqlStr + " 	, sum(l.jungsanPrice) as totSumJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 100 then l.jungsanPrice else 0 end) as sumCardJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 20 then l.jungsanPrice else 0 end) as sumBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 7 then l.jungsanPrice else 0 end) as sumVBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 77 then l.jungsanPrice else 0 end) as sumTenOutBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 6 then l.jungsanPrice else 0 end) as sumTenInBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 400 then l.jungsanPrice else 0 end) as sumHPJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 550 then l.jungsanPrice else 0 end) as sumGifttingJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 560 then l.jungsanPrice else 0 end) as sumGifticonJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 110 then l.jungsanPrice else 0 end) as sumOKJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 80 then l.jungsanPrice else 0 end) as sumAllAtJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 150 then l.jungsanPrice else 0 end) as sumRentalJungsanPrice "

		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen3' then l.jungsanPrice else 0 end) as sumteenxteen3JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen4' then l.jungsanPrice else 0 end) as sumteenxteen4JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen5' then l.jungsanPrice else 0 end) as sumteenxteen5JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen6' then l.jungsanPrice else 0 end) as sumteenxteen6JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen8' then l.jungsanPrice else 0 end) as sumteenxteen8JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen9' then l.jungsanPrice else 0 end) as sumteenxteen9JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteen10' then l.jungsanPrice else 0 end) as sumteenteen10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten01' then l.jungsanPrice else 0 end) as sumtenbyten01JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten02' then l.jungsanPrice else 0 end) as sumtenbyten02JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteeha' then l.jungsanPrice else 0 end) as sumteenxteehaJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteenr' then l.jungsanPrice else 0 end) as sumteenxteenrJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteensp' then l.jungsanPrice else 0 end) as sumteenteenspJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteenap' then l.jungsanPrice else 0 end) as sumteenteenapJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'KCTEN0001m' then l.jungsanPrice else 0 end) as sumKCTEN0001mJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'newkakaopay' then l.jungsanPrice else 0 end) as sumKakaoJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when (l.PGuserid = 'naverpay' and l.appMethod<>80) then l.jungsanPrice else 0 end) as sumnaverpayJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when (l.PGuserid = 'naverpay' and l.appMethod=80) then l.jungsanPrice else 0 end) as sumnaverpayJungsanPoint "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'payco' then l.jungsanPrice else 0 end) as sumpaycoJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum' then l.jungsanPrice else 0 end) as sumbankipkumJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_10x10' then l.jungsanPrice else 0 end) as sumbankipkum_10x10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_fingers' then l.jungsanPrice else 0 end) as sumbankipkum_fingersJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund' then l.jungsanPrice else 0 end) as sumbankrefundJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_10x10' then l.jungsanPrice else 0 end) as sumbankrefund_10x10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_fingers' then l.jungsanPrice else 0 end) as sumbankrefund_fingersJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = '10x10_2' then l.jungsanPrice else 0 end) as sum10x10_2JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'R5523' then l.jungsanPrice else 0 end) as sumR5523JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'mobilians' then l.jungsanPrice else 0 end) as summobiliansJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'gifticon' then l.jungsanPrice else 0 end) as sumPGgifticonJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'giftting' then l.jungsanPrice else 0 end) as sumPGgifttingJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'okcashbag' then l.jungsanPrice else 0 end) as sumPGokcashbagJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'toss' then l.jungsanPrice else 0 end) as sumPGtossJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'chai' then l.jungsanPrice else 0 end) as sumPGchaiJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'convinienspay' then l.jungsanPrice else 0 end) as sumPGConvinienspayJungsanPrice "

		sqlStr = sqlStr + " from T_LIST l"
        sqlStr = sqlStr + " group by yyyymmdd"
        sqlStr = sqlStr + " order by yyyymmdd"

		''response.write sqlStr & "<Br>"
		''response.end
		'rsget.Open sqlStr,dbget,1
		''방식 수정 2016/03/31 by eastone
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalcount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CPGDataStatisticItem

				FItemList(i).Fyyyymmdd = rsget("yyyymmdd")

				FItemList(i).FtotSumPrice = rsget("totSumPrice")

				FItemList(i).FsumCardPrice = rsget("sumCardPrice")
				FItemList(i).FsumBankPrice = rsget("sumBankPrice")
				FItemList(i).FsumVBankPrice = rsget("sumVBankPrice")
				FItemList(i).FsumTenOutBankPrice = rsget("sumTenOutBankPrice")
				FItemList(i).FsumTenInBankPrice = rsget("sumTenInBankPrice")
				FItemList(i).FsumHPPrice = rsget("sumHPPrice")
				FItemList(i).FsumGifttingPrice = rsget("sumGifttingPrice")
				FItemList(i).FsumGifticonPrice = rsget("sumGifticonPrice")
				FItemList(i).FsumOKPrice = rsget("sumOKPrice")
				FItemList(i).FsumAllAtPrice = rsget("sumAllAtPrice")

				FItemList(i).Fsumteenxteen3Price = rsget("sumteenxteen3Price")
				FItemList(i).Fsumteenxteen4Price = rsget("sumteenxteen4Price")
				FItemList(i).Fsumteenxteen5Price = rsget("sumteenxteen5Price")
				FItemList(i).Fsumteenxteen6Price = rsget("sumteenxteen6Price")
				FItemList(i).Fsumteenxteen8Price = rsget("sumteenxteen8Price")
				FItemList(i).Fsumteenxteen9Price = rsget("sumteenxteen9Price")
				FItemList(i).Fsumteenteen10Price = rsget("sumteenteen10Price")
				FItemList(i).Fsumtenbyten01Price = rsget("sumtenbyten01Price")
				FItemList(i).Fsumtenbyten02Price = rsget("sumtenbyten02Price")
				FItemList(i).FsumteenxteehaPrice = rsget("sumteenxteehaPrice")
				FItemList(i).FsumteenxteenrPrice = rsget("sumteenxteenrPrice")
				FItemList(i).FsumteenteenspPrice = rsget("sumteenteenspPrice")
				FItemList(i).FsumteenteenapPrice = rsget("sumteenteenapPrice")
				FItemList(i).FsumKCTEN0001mPrice = rsget("sumKCTEN0001mPrice")
				FItemList(i).FsumKakaopayPrice = rsget("sumKakaopayPrice")
				FItemList(i).FsumnaverpayPrice = rsget("sumnaverpayPrice")
				FItemList(i).FsumnaverpayPoint = rsget("sumnaverpayPoint")
				FItemList(i).FsumpaycoPrice = rsget("sumpaycoPrice")
				FItemList(i).FsumbankipkumPrice = rsget("sumbankipkumPrice")
				FItemList(i).Fsumbankipkum_10x10Price = rsget("sumbankipkum_10x10Price")
				FItemList(i).Fsumbankipkum_fingersPrice = rsget("sumbankipkum_fingersPrice")
				FItemList(i).FsumbankrefundPrice = rsget("sumbankrefundPrice")
				FItemList(i).Fsumbankrefund_10x10Price = rsget("sumbankrefund_10x10Price")
				FItemList(i).Fsumbankrefund_fingersPrice = rsget("sumbankrefund_fingersPrice")
				FItemList(i).Fsum10x10_2Price = rsget("sum10x10_2Price")
				FItemList(i).FsumR5523Price = rsget("sumR5523Price")
				FItemList(i).FsummobiliansPrice = rsget("summobiliansPrice")
				FItemList(i).FsumPGgifticonPrice = rsget("sumPGgifticonPrice")
				FItemList(i).FsumPGgifttingPrice = rsget("sumPGgifttingPrice")
				FItemList(i).FsumPGokcashbagPrice = rsget("sumPGokcashbagPrice")
				FItemList(i).FsumPGtossPrice = rsget("sumPGtossPrice")
				FItemList(i).FsumPGchaiPrice = rsget("sumPGchaiPrice")
				FItemList(i).FsumPGConvinienspayPrice = rsget("sumPGConvinienspayPrice")

				FItemList(i).FtotSumJungsanPrice = rsget("totSumJungsanPrice")

				FItemList(i).FsumCardJungsanPrice = rsget("sumCardJungsanPrice")
				FItemList(i).FsumBankJungsanPrice = rsget("sumBankJungsanPrice")
				FItemList(i).FsumVBankJungsanPrice = rsget("sumVBankJungsanPrice")
				FItemList(i).FsumTenOutBankJungsanPrice = rsget("sumTenOutBankJungsanPrice")
				FItemList(i).FsumTenInBankJungsanPrice = rsget("sumTenInBankJungsanPrice")
				FItemList(i).FsumHPJungsanPrice = rsget("sumHPJungsanPrice")
				FItemList(i).FsumGifttingJungsanPrice = rsget("sumGifttingJungsanPrice")
				FItemList(i).FsumGifticonJungsanPrice = rsget("sumGifticonJungsanPrice")
				FItemList(i).FsumOKJungsanPrice = rsget("sumOKJungsanPrice")
				FItemList(i).FsumAllAtJungsanPrice = rsget("sumAllAtJungsanPrice")

				FItemList(i).Fsumteenxteen3JungsanPrice = rsget("sumteenxteen3JungsanPrice")
				FItemList(i).Fsumteenxteen4JungsanPrice = rsget("sumteenxteen4JungsanPrice")
				FItemList(i).Fsumteenxteen5JungsanPrice = rsget("sumteenxteen5JungsanPrice")
				FItemList(i).Fsumteenxteen6JungsanPrice = rsget("sumteenxteen6JungsanPrice")
				FItemList(i).Fsumteenxteen8JungsanPrice = rsget("sumteenxteen8JungsanPrice")
				FItemList(i).Fsumteenxteen9JungsanPrice = rsget("sumteenxteen9JungsanPrice")
				FItemList(i).Fsumteenteen10JungsanPrice = rsget("sumteenteen10JungsanPrice")
				FItemList(i).Fsumtenbyten01JungsanPrice = rsget("sumtenbyten01JungsanPrice")
				FItemList(i).Fsumtenbyten02JungsanPrice = rsget("sumtenbyten02JungsanPrice")
				FItemList(i).FsumteenxteehaJungsanPrice = rsget("sumteenxteehaJungsanPrice")
				FItemList(i).FsumteenxteenrJungsanPrice = rsget("sumteenxteenrJungsanPrice")
				FItemList(i).FsumteenteenspJungsanPrice = rsget("sumteenteenspJungsanPrice")
				FItemList(i).FsumteenteenapJungsanPrice = rsget("sumteenteenapJungsanPrice")
				FItemList(i).FsumKCTEN0001mJungsanPrice = rsget("sumKCTEN0001mJungsanPrice")
				FItemList(i).FsumKakaoJungsanPrice = rsget("sumKakaoJungsanPrice")
				FItemList(i).FsumnaverpayJungsanPrice = rsget("sumnaverpayJungsanPrice")
				FItemList(i).FsumnaverpayJungsanPoint = rsget("sumnaverpayJungsanPoint")
				FItemList(i).FsumpaycoJungsanPrice = rsget("sumpaycoJungsanPrice")
				FItemList(i).FsumbankipkumJungsanPrice = rsget("sumbankipkumJungsanPrice")
				FItemList(i).Fsumbankipkum_10x10JungsanPrice = rsget("sumbankipkum_10x10JungsanPrice")
				FItemList(i).Fsumbankipkum_fingersJungsanPrice = rsget("sumbankipkum_fingersJungsanPrice")
				FItemList(i).FsumbankrefundJungsanPrice = rsget("sumbankrefundJungsanPrice")
				FItemList(i).Fsumbankrefund_10x10JungsanPrice = rsget("sumbankrefund_10x10JungsanPrice")
				FItemList(i).Fsumbankrefund_fingersJungsanPrice = rsget("sumbankrefund_fingersJungsanPrice")
				FItemList(i).Fsum10x10_2JungsanPrice = rsget("sum10x10_2JungsanPrice")
				FItemList(i).FsumR5523JungsanPrice = rsget("sumR5523JungsanPrice")
				FItemList(i).FsummobiliansJungsanPrice = rsget("summobiliansJungsanPrice")
				FItemList(i).FsumPGgifticonJungsanPrice = rsget("sumPGgifticonJungsanPrice")
				FItemList(i).FsumPGgifttingJungsanPrice = rsget("sumPGgifttingJungsanPrice")
				FItemList(i).FsumPGokcashbagJungsanPrice = rsget("sumPGokcashbagJungsanPrice")
				FItemList(i).FsumPGtossJungsanPrice = rsget("sumPGtossJungsanPrice")
				FItemList(i).FsumPGchaiJungsanPrice = rsget("sumPGchaiJungsanPrice")
				FItemList(i).FsumPGConvinienspayJungsanPrice = rsget("sumPGConvinienspayJungsanPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function getPGDataStatisticList_ON_old_20160331()
		dim i, j, sqlStr, addSqlStr

		'// ====================================================================
		sqlStr = " select "

		if (FRectDateGubun = "ipkumdate") then
			sqlStr = sqlStr + " 	l.ipkumdate AS yyyymmdd "
		else
			sqlStr = sqlStr + " 	convert(VARCHAR(10), IsNull(l.cancelDate, l.appDate), 127) AS yyyymmdd "
		end if

		sqlStr = sqlStr + " 	, sum(l.appPrice) as totSumPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 100 then l.appPrice else 0 end) as sumCardPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 20 then l.appPrice else 0 end) as sumBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 7 then l.appPrice else 0 end) as sumVBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 77 then l.appPrice else 0 end) as sumTenOutBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 6 then l.appPrice else 0 end) as sumTenInBankPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 400 then l.appPrice else 0 end) as sumHPPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 550 then l.appPrice else 0 end) as sumGifttingPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 560 then l.appPrice else 0 end) as sumGifticonPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 110 then l.appPrice else 0 end) as sumOKPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 80 then l.appPrice else 0 end) as sumAllAtPrice "

		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen3' then l.appPrice else 0 end) as sumteenxteen3Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen4' then l.appPrice else 0 end) as sumteenxteen4Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen5' then l.appPrice else 0 end) as sumteenxteen5Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen6' then l.appPrice else 0 end) as sumteenxteen6Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen8' then l.appPrice else 0 end) as sumteenxteen8Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen9' then l.appPrice else 0 end) as sumteenxteen9Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteen10' then l.appPrice else 0 end) as sumteenteen10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten01' then l.appPrice else 0 end) as sumtenbyten01Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten02' then l.appPrice else 0 end) as sumtenbyten02Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'KCTEN0001m' then l.appPrice else 0 end) as sumKCTEN0001mPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum' then l.appPrice else 0 end) as sumbankipkumPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_10x10' then l.appPrice else 0 end) as sumbankipkum_10x10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_fingers' then l.appPrice else 0 end) as sumbankipkum_fingersPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund' then l.appPrice else 0 end) as sumbankrefundPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_10x10' then l.appPrice else 0 end) as sumbankrefund_10x10Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_fingers' then l.appPrice else 0 end) as sumbankrefund_fingersPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = '10x10_2' then l.appPrice else 0 end) as sum10x10_2Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'R5523' then l.appPrice else 0 end) as sumR5523Price "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'mobilians' then l.appPrice else 0 end) as summobiliansPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'gifticon' then l.appPrice else 0 end) as sumPGgifticonPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'giftting' then l.appPrice else 0 end) as sumPGgifttingPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'okcashbag' then l.appPrice else 0 end) as sumPGokcashbagPrice "

		sqlStr = sqlStr + " 	, sum(l.jungsanPrice) as totSumJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 100 then l.jungsanPrice else 0 end) as sumCardJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 20 then l.jungsanPrice else 0 end) as sumBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 7 then l.jungsanPrice else 0 end) as sumVBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 77 then l.jungsanPrice else 0 end) as sumTenOutBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 6 then l.jungsanPrice else 0 end) as sumTenInBankJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 400 then l.jungsanPrice else 0 end) as sumHPJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 550 then l.jungsanPrice else 0 end) as sumGifttingJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 560 then l.jungsanPrice else 0 end) as sumGifticonJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 110 then l.jungsanPrice else 0 end) as sumOKJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.appMethod = 80 then l.jungsanPrice else 0 end) as sumAllAtJungsanPrice "

		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen3' then l.jungsanPrice else 0 end) as sumteenxteen3JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen4' then l.jungsanPrice else 0 end) as sumteenxteen4JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen5' then l.jungsanPrice else 0 end) as sumteenxteen5JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen6' then l.jungsanPrice else 0 end) as sumteenxteen6JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen8' then l.jungsanPrice else 0 end) as sumteenxteen8JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenxteen9' then l.jungsanPrice else 0 end) as sumteenxteen9JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'teenteen10' then l.jungsanPrice else 0 end) as sumteenteen10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten01' then l.jungsanPrice else 0 end) as sumtenbyten01JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'tenbyten02' then l.jungsanPrice else 0 end) as sumtenbyten02JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'KCTEN0001m' then l.jungsanPrice else 0 end) as sumKCTEN0001mJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum' then l.jungsanPrice else 0 end) as sumbankipkumJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_10x10' then l.jungsanPrice else 0 end) as sumbankipkum_10x10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankipkum_fingers' then l.jungsanPrice else 0 end) as sumbankipkum_fingersJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund' then l.jungsanPrice else 0 end) as sumbankrefundJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_10x10' then l.jungsanPrice else 0 end) as sumbankrefund_10x10JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'bankrefund_fingers' then l.jungsanPrice else 0 end) as sumbankrefund_fingersJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = '10x10_2' then l.jungsanPrice else 0 end) as sum10x10_2JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'R5523' then l.jungsanPrice else 0 end) as sumR5523JungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'mobilians' then l.jungsanPrice else 0 end) as summobiliansJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'gifticon' then l.jungsanPrice else 0 end) as sumPGgifticonJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'giftting' then l.jungsanPrice else 0 end) as sumPGgifttingJungsanPrice "
		sqlStr = sqlStr + " 	, sum(case when l.PGuserid = 'okcashbag' then l.jungsanPrice else 0 end) as sumPGokcashbagJungsanPrice "

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		if (FRectStartIpkumdate <> "") then
			sqlStr = sqlStr + " 	and l.ipkumdate >= '" + CStr(FRectStartIpkumdate) + "' "
			sqlStr = sqlStr + " 	and l.ipkumdate < '" + CStr(FRectEndIpkumDate) + "' "
		end if

		if (FRectStartdate <> "") then
			sqlStr = sqlStr + " 	and IsNull(l.cancelDate, l.appDate) >= '" + CStr(FRectStartdate) + "' "
			sqlStr = sqlStr + " 	and IsNull(l.cancelDate, l.appDate) < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectPGuserid <> "") then
			sqlStr = sqlStr + " 	and l.PGuserid = '" + CStr(FRectPGuserid) + "' "
		end if

		if (FRectSiteName <> "") then
			if (FRectSiteName = "10x10all") then
				sqlStr = sqlStr + " and l.sitename in ('10x10', '10x10mobile') "
			else
				sqlStr = sqlStr + " and l.sitename = '" + CStr(FRectSiteName) + "' "
			end if
		end if

	    if (FRectPGGubun <> "") then
    	    sqlStr = sqlStr + " and l.pggubun = '" + CStr(FRectPGGubun) + "' "
    	end if

		if (FRectReasonGubun <> "") then
			if (FRectReasonGubun = "XXX") then
				sqlStr = sqlStr + " and IsNull(l.reasonGubun, '') not in ('001', '002', '003', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') "
			else
				sqlStr = sqlStr + " and IsNull(l.reasonGubun, '') = '" + CStr(FRectReasonGubun) + "' "
			end if
		end if

		if (FRectDateGubun = "ipkumdate") then
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	l.ipkumdate "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	l.ipkumdate "
		else
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	convert(VARCHAR(10), IsNull(l.cancelDate, l.appDate), 127) "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	convert(VARCHAR(10), IsNull(l.cancelDate, l.appDate), 127) "
		end if

		''response.write sqlStr & "<Br>"
		'rsget.Open sqlStr,dbget,1
		''방식 수정 2016/03/31 by eastone
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalcount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CPGDataStatisticItem

				FItemList(i).Fyyyymmdd = rsget("yyyymmdd")

				FItemList(i).FtotSumPrice = rsget("totSumPrice")

				FItemList(i).FsumCardPrice = rsget("sumCardPrice")
				FItemList(i).FsumBankPrice = rsget("sumBankPrice")
				FItemList(i).FsumVBankPrice = rsget("sumVBankPrice")
				FItemList(i).FsumTenOutBankPrice = rsget("sumTenOutBankPrice")
				FItemList(i).FsumTenInBankPrice = rsget("sumTenInBankPrice")
				FItemList(i).FsumHPPrice = rsget("sumHPPrice")
				FItemList(i).FsumGifttingPrice = rsget("sumGifttingPrice")
				FItemList(i).FsumGifticonPrice = rsget("sumGifticonPrice")
				FItemList(i).FsumOKPrice = rsget("sumOKPrice")
				FItemList(i).FsumAllAtPrice = rsget("sumAllAtPrice")

				FItemList(i).Fsumteenxteen3Price = rsget("sumteenxteen3Price")
				FItemList(i).Fsumteenxteen4Price = rsget("sumteenxteen4Price")
				FItemList(i).Fsumteenxteen5Price = rsget("sumteenxteen5Price")
				FItemList(i).Fsumteenxteen6Price = rsget("sumteenxteen6Price")
				FItemList(i).Fsumteenxteen8Price = rsget("sumteenxteen8Price")
				FItemList(i).Fsumteenxteen9Price = rsget("sumteenxteen9Price")
				FItemList(i).Fsumteenteen10Price = rsget("sumteenteen10Price")
				FItemList(i).Fsumtenbyten01Price = rsget("sumtenbyten01Price")
				FItemList(i).Fsumtenbyten02Price = rsget("sumtenbyten02Price")
				FItemList(i).FsumKCTEN0001mPrice = rsget("sumKCTEN0001mPrice")
				FItemList(i).FsumbankipkumPrice = rsget("sumbankipkumPrice")
				FItemList(i).Fsumbankipkum_10x10Price = rsget("sumbankipkum_10x10Price")
				FItemList(i).Fsumbankipkum_fingersPrice = rsget("sumbankipkum_fingersPrice")
				FItemList(i).FsumbankrefundPrice = rsget("sumbankrefundPrice")
				FItemList(i).Fsumbankrefund_10x10Price = rsget("sumbankrefund_10x10Price")
				FItemList(i).Fsumbankrefund_fingersPrice = rsget("sumbankrefund_fingersPrice")
				FItemList(i).Fsum10x10_2Price = rsget("sum10x10_2Price")
				FItemList(i).FsumR5523Price = rsget("sumR5523Price")
				FItemList(i).FsummobiliansPrice = rsget("summobiliansPrice")
				FItemList(i).FsumPGgifticonPrice = rsget("sumPGgifticonPrice")
				FItemList(i).FsumPGgifttingPrice = rsget("sumPGgifttingPrice")
				FItemList(i).FsumPGokcashbagPrice = rsget("sumPGokcashbagPrice")

				FItemList(i).FtotSumJungsanPrice = rsget("totSumJungsanPrice")

				FItemList(i).FsumCardJungsanPrice = rsget("sumCardJungsanPrice")
				FItemList(i).FsumBankJungsanPrice = rsget("sumBankJungsanPrice")
				FItemList(i).FsumVBankJungsanPrice = rsget("sumVBankJungsanPrice")
				FItemList(i).FsumTenOutBankJungsanPrice = rsget("sumTenOutBankJungsanPrice")
				FItemList(i).FsumTenInBankJungsanPrice = rsget("sumTenInBankJungsanPrice")
				FItemList(i).FsumHPJungsanPrice = rsget("sumHPJungsanPrice")
				FItemList(i).FsumGifttingJungsanPrice = rsget("sumGifttingJungsanPrice")
				FItemList(i).FsumGifticonJungsanPrice = rsget("sumGifticonJungsanPrice")
				FItemList(i).FsumOKJungsanPrice = rsget("sumOKJungsanPrice")
				FItemList(i).FsumAllAtJungsanPrice = rsget("sumAllAtJungsanPrice")

				FItemList(i).Fsumteenxteen3JungsanPrice = rsget("sumteenxteen3JungsanPrice")
				FItemList(i).Fsumteenxteen4JungsanPrice = rsget("sumteenxteen4JungsanPrice")
				FItemList(i).Fsumteenxteen5JungsanPrice = rsget("sumteenxteen5JungsanPrice")
				FItemList(i).Fsumteenxteen6JungsanPrice = rsget("sumteenxteen6JungsanPrice")
				FItemList(i).Fsumteenxteen8JungsanPrice = rsget("sumteenxteen8JungsanPrice")
				FItemList(i).Fsumteenxteen9JungsanPrice = rsget("sumteenxteen9JungsanPrice")
				FItemList(i).Fsumteenteen10JungsanPrice = rsget("sumteenteen10JungsanPrice")
				FItemList(i).Fsumtenbyten01JungsanPrice = rsget("sumtenbyten01JungsanPrice")
				FItemList(i).Fsumtenbyten02JungsanPrice = rsget("sumtenbyten02JungsanPrice")
				FItemList(i).FsumKCTEN0001mJungsanPrice = rsget("sumKCTEN0001mJungsanPrice")
				FItemList(i).FsumbankipkumJungsanPrice = rsget("sumbankipkumJungsanPrice")
				FItemList(i).Fsumbankipkum_10x10JungsanPrice = rsget("sumbankipkum_10x10JungsanPrice")
				FItemList(i).Fsumbankipkum_fingersJungsanPrice = rsget("sumbankipkum_fingersJungsanPrice")
				FItemList(i).FsumbankrefundJungsanPrice = rsget("sumbankrefundJungsanPrice")
				FItemList(i).Fsumbankrefund_10x10JungsanPrice = rsget("sumbankrefund_10x10JungsanPrice")
				FItemList(i).Fsumbankrefund_fingersJungsanPrice = rsget("sumbankrefund_fingersJungsanPrice")
				FItemList(i).Fsum10x10_2JungsanPrice = rsget("sum10x10_2JungsanPrice")
				FItemList(i).FsumR5523JungsanPrice = rsget("sumR5523JungsanPrice")
				FItemList(i).FsummobiliansJungsanPrice = rsget("summobiliansJungsanPrice")
				FItemList(i).FsumPGgifticonJungsanPrice = rsget("sumPGgifticonJungsanPrice")
				FItemList(i).FsumPGgifttingJungsanPrice = rsget("sumPGgifttingJungsanPrice")
				FItemList(i).FsumPGokcashbagJungsanPrice = rsget("sumPGokcashbagJungsanPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function GetArrCardComp()
		GetArrCardComp = Split(FArrCardComp, "|")
	end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

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
