<%
'###########################################################
' Description : 기프트카드 주문 클래스
' Hieditor : 2009.04.17 이상구 생성
'			 2016.07.20 한용민 수정
'###########################################################

'프런트와 어드민상의 cs_giftcard_ordercls.asp 는 동일한 내용이다.
'만약 내용이 틀리면 어드민상의 cs_giftcard_ordercls.asp 를 복사해온다.
'프런트에 내용을 추가하면 어드민상의 클래스도 추가해준다.

Class cGiftcardOrderItem
	public FgiftOrderSerial
	public Fidx
	public FcardItemid
	public FcardOption
	public FmasterCardCode
	public FmasterCheckCode
	public Fuserid
	public Fbuyname
	public Fbuyemail
	public Fbuyhp
	public FbuyPhone
	public Fsendhp
	public Fsendemail
	public Freqhp
	public Freqemail
	public Ftotalsum
	public Fjumundiv
	public Faccountdiv
	public Fipkumname
	public Faccountno
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fcancelyn
	public Fpaydateid
	public Fresultmsg
	public Fauthcode
	public Fdiscountrate
	public Fsubtotalprice
	public Fmiletotalprice
	public Ftencardspend
	public Fcashreceiptreq
	public FinsureCd
	public FinsureMsg
	public Freferip
	public Fuserlevel
	public Fcanceldate
	public FsumPaymentEtc
	public FsendDiv
	public FbookingYn
	public FbookingDate
	public FsendDate
	public FdesignId
	public FMMSTitle
	public FMMSContent
	public FemailTitle
	public FemailContent

	public FSmallimage
	public FCarditemname
	public FCardinfo
	public FCarddesc
	public FcardPrice
	public FcardregDate
	public Freguserid
	public FcardStatus
	public FcardOptionName

    public FcpnRegUserID
	public FcouponNo

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

    function GetAccountdivName()
        dim oacctdiv
        if IsNULL(FAccountdiv) then Exit function
        oacctdiv = Trim(FAccountdiv)

        select case oacctdiv
            case "7"
                : GetAccountdivName = "무통장"
            case "14"
                : GetAccountdivName = "편의점결제"
            case "100"
                : GetAccountdivName = "신용카드"
            case "20"
                : GetAccountdivName = "실시간계좌이체"
            case "80"
                : GetAccountdivName = "All@멤버쉽카드"
            case "50"
                : GetAccountdivName = "외부몰결제"
            case "30"
                : GetAccountdivName = "포인트"
            case "90"
                : GetAccountdivName = "상품권"
            case "110"
                : GetAccountdivName = "신용카드+OK캐쉬백"
            case "400"
                : GetAccountdivName = "핸드폰결제"
            case else
                : GetAccountdivName = ""
        end select
    end function

    function GetIpkumDivName()
        dim oipkumdiv
        if IsNULL(Fipkumdiv) then Exit function
        oipkumdiv = Trim(Fipkumdiv)

        select case oipkumdiv
            case "0"
                : GetIpkumDivName = "주문실패"
            case "1"
                : GetIpkumDivName = "주문실패"
            case "2"
                : GetIpkumDivName = "주문접수"
            case "3"
                : GetIpkumDivName = "입금대기"
            case "4"
                : GetIpkumDivName = "결제완료"
            case "5"
                : GetIpkumDivName = "전송대기"
            case "8"
                : GetIpkumDivName = "전송완료"
            case "9"
                : GetIpkumDivName = "주문취소"
            case else
                : GetIpkumDivName = ""
        end select
    end function

    function GetCardStatusName()
        dim ocardStatus
        if IsNULL(FcardStatus) then FcardStatus = "0"
        ocardStatus = Trim(FcardStatus)

        select case ocardStatus
            case "1"
                : GetCardStatusName = "등록완료"
            case "3"
                : GetCardStatusName = "등록취소"
            case "5"
                : GetCardStatusName = "카드만료"
            case "0"
                : GetCardStatusName = "등록이전"
            case else
                : GetCardStatusName = ""
        end select
    end function

	public function GetCardStatusColor()
		if FcardStatus="0" then
			GetCardStatusColor="#FF0000"
		elseif FcardStatus="1" then
			GetCardStatusColor="#44BBBB"
		elseif FcardStatus="2" then
			GetCardStatusColor="#000000"
		elseif FcardStatus="3" then
			GetCardStatusColor="#000000"
		elseif FcardStatus="4" then
			GetCardStatusColor="#0000FF"
		elseif FcardStatus="5" then
			GetCardStatusColor="#CC9933"
		elseif FcardStatus="6" then
			GetCardStatusColor="#FF00FF"
		elseif FcardStatus="7" then
			GetCardStatusColor="#EE2222"
		elseif FcardStatus="8" then
			GetCardStatusColor="#EE2222"
		elseif FcardStatus="9" then
			GetCardStatusColor="#FF0000"
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

	public function GetJumunDivColor()
		if Fjumundiv="0" then
			GetJumunDivColor="#FF0000"
		elseif Fjumundiv="1" then
			GetJumunDivColor="#44BBBB"
		elseif Fjumundiv="2" then
			GetJumunDivColor="#000000"
		elseif Fjumundiv="3" then
			GetJumunDivColor="#000000"
		elseif Fjumundiv="4" then
			GetJumunDivColor="#0000FF"
		elseif Fjumundiv="5" then
			GetJumunDivColor="#CC9933"
		elseif Fjumundiv="6" then
			GetJumunDivColor="#FF00FF"
		elseif Fjumundiv="7" then
			GetJumunDivColor="#EE2222"
		elseif Fjumundiv="8" then
			GetJumunDivColor="#EE2222"
		elseif Fjumundiv="9" then
			GetJumunDivColor="#FF0000"
		end if
	end function

    function GetJumunDivName()
        dim ojumundiv
        if IsNULL(Fjumundiv) then Exit function
        ojumundiv = Trim(Fjumundiv)

        select case ojumundiv
            case "1"
                : GetJumunDivName = "결제대기"
            case "3"
                : GetJumunDivName = "전송대기"
            case "5"
                : GetJumunDivName = "전송완료"
            case "7"
                : GetJumunDivName = "등록완료"
            case "9"
                : GetJumunDivName = "주문취소"
            case else
                : GetJumunDivName = ""
        end select
    end function

	'/사용금지		'/공용펑션에 공용함수 쓸것		'/2016.07.20 한용민
	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44DD44"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444FF"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#FF1111"   ''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#FF0000"   ''VIP GOLD
		elseif Fuserlevel="9" then
			GetUserLevelColor = "#FF11FF"  '' mania
		elseif Fuserlevel="7" then
			GetUserLevelColor = "#FF11FF"  '' staff
		elseif Fuserlevel="6" then
			GetUserLevelColor = "#FF11FF"  '' friends
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

	'/사용금지		'/공용펑션에 공용함수 쓸것		'/2016.07.20 한용민
	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Green"
		elseif Fuserlevel="2" then
			GetUserLevelName = "Blue"
		elseif Fuserlevel="3" then
			GetUserLevelName = "VIP Silver"
		elseif Fuserlevel="4" then
			GetUserLevelName = "VIP Gold"
	    elseif Fuserlevel="7" then
			GetUserLevelName = "Staff"
		elseif Fuserlevel="9" then
			GetUserLevelName = "Mania"  ''magenta
		else
			GetUserLevelName = "Yellow"
		end if
	end function

    ''주문취소 (웹 취소가능)
    public function IsWebOrderCancelEnable()
        IsWebOrderCancelEnable = false
        if (Not IsValidOrder) then Exit function

        IsWebOrderCancelEnable = (Fjumundiv<7)

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

    public function IsValidOrder()
        IsValidOrder = (FIpkumdiv>1) and (FCancelyn="N")
    end function

    ''=================================================================================================
    '' 각종 증명서 관련 : 상품권은 증빙서류가 없다.(계산서X, 현금영수증X)

    ''전자보증서 존재
    public function IsInsureDocExists()
        IsInsureDocExists = (FInsureCd="0")
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


Class cGiftCardOrder
	public FItemList()
	public FOneItem

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FGiftCardWebImgUrl

	public FRectGiftOrderSerial
	public FRectUserID
	public FRectBuyname
	public FRectIpkumName

	public FRectBuyHp
	public FRectReqHp
	public FRectBuyPhone

	public FRectRegStart
	public FRectRegEnd
	public frectcpnreguserid

	public Sub getCSGiftcardOrderList()
		dim sqlStr, i

		sqlStr = "EXEC [db_order].[dbo].[usp_cs_GiftCard_orderList_Cnt] '" + CStr(FRectGiftOrderSerial) + "', '" + CStr(FRectUserID) + "', '" + CStr(FRectBuyname) + "', '" + CStr(FRectIpkumName) + "', '" + CStr(FRectBuyHp) + "', '" + CStr(FRectReqHp) + "', '" + CStr(FRectBuyPhone) + "', '" + CStr(FRectRegStart) + "', '" + CStr(FRectRegEnd) + "', N'" + CStr(frectcpnreguserid) + "'"
	    rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    If Not rsget.Eof Then
		    FTotalCount = rsget("cnt")
		End If
		rsget.close

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		sqlStr = "EXEC [db_order].[dbo].[usp_cs_GiftCard_orderList] " +CStr(FPageSize) + ", " +CStr(FCurrPage) + ", '" + CStr(FRectGiftOrderSerial) + "', '" + CStr(FRectUserID) + "', '" + CStr(FRectBuyname) + "', '" + CStr(FRectIpkumName) + "', '" + CStr(FRectBuyHp) + "', '" + CStr(FRectReqHp) + "', '" + CStr(FRectBuyPhone) + "', '" + CStr(FRectRegStart) + "', '" + CStr(FRectRegEnd) + "', N'" + CStr(frectcpnreguserid) + "'"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        FResultcount = rsget.Recordcount

		redim preserve FItemList(FResultCount)

        i = 0
        if Not rsget.Eof then
			do until rsget.Eof
				set FItemList(i) = new cGiftcardOrderItem

				FItemList(i).FgiftOrderSerial	= rsget("giftOrderSerial")
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).FcardItemid		= rsget("cardItemid")
				FItemList(i).FcardOption		= rsget("cardOption")
				FItemList(i).FmasterCardCode	= rsget("masterCardCode")
				FItemList(i).FmasterCheckCode	= rsget("masterCheckCode")
				FItemList(i).Fuserid			= rsget("userid")
				FItemList(i).Fbuyname			= rsget("buyname")
				FItemList(i).Fbuyemail			= rsget("buyemail")
				FItemList(i).Fbuyhp				= rsget("buyhp")
				FItemList(i).FbuyPhone			= rsget("buyPhone")
				FItemList(i).Fsendhp			= rsget("sendhp")
				FItemList(i).Fsendemail			= rsget("sendemail")
				FItemList(i).Freqhp				= rsget("reqhp")
				FItemList(i).Freqemail			= rsget("reqemail")
				FItemList(i).Ftotalsum			= rsget("totalsum")
				FItemList(i).Fjumundiv			= rsget("jumundiv")
				FItemList(i).Faccountdiv		= rsget("accountdiv")
				FItemList(i).Fipkumname			= rsget("ipkumname")
				FItemList(i).Faccountno			= rsget("accountno")
				FItemList(i).Fipkumdiv			= rsget("ipkumdiv")
				FItemList(i).Fipkumdate			= rsget("ipkumdate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fcancelyn			= rsget("cancelyn")
				FItemList(i).Fpaydateid			= rsget("paydateid")
				FItemList(i).Fresultmsg			= rsget("resultmsg")
				FItemList(i).Fauthcode			= rsget("authcode")
				FItemList(i).Fdiscountrate		= rsget("discountrate")
				FItemList(i).Fsubtotalprice		= rsget("subtotalprice")
				FItemList(i).Fmiletotalprice	= rsget("miletotalprice")
				FItemList(i).Ftencardspend		= rsget("tencardspend")
				FItemList(i).Fcashreceiptreq	= rsget("cashreceiptreq")
				FItemList(i).FinsureCd			= rsget("insureCd")
				FItemList(i).FinsureMsg			= rsget("insureMsg")
				FItemList(i).Freferip			= rsget("referip")
				FItemList(i).Fuserlevel			= rsget("userlevel")
				FItemList(i).Fcanceldate		= rsget("canceldate")
				FItemList(i).FsumPaymentEtc		= rsget("sumPaymentEtc")
				FItemList(i).FsendDiv			= rsget("sendDiv")
				FItemList(i).FbookingYn			= rsget("bookingYn")
				FItemList(i).FbookingDate		= rsget("bookingDate")
				FItemList(i).FsendDate			= rsget("sendDate")
				FItemList(i).FdesignId			= rsget("designId")
				FItemList(i).FMMSTitle			= db2html(rsget("MMSTitle"))
				FItemList(i).FMMSContent		= db2html(rsget("MMSContent"))
				FItemList(i).FemailTitle		= db2html(rsget("emailTitle"))
				FItemList(i).FemailContent		= db2html(rsget("emailContent"))

				FItemList(i).FSmallimage		= FGiftCardWebImgUrl & "/giftcard/small/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("smallImage")
				FItemList(i).FCarditemname		= rsget("Carditemname")
				FItemList(i).FCardinfo			= rsget("Cardinfo")
				FItemList(i).FCarddesc			= rsget("Carddesc")
				FItemList(i).FcardPrice			= rsget("cardPrice")
				FItemList(i).FcardregDate		= rsget("cardregDate")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).FcardStatus		= rsget("cardStatus")
				FItemList(i).FcardOptionName	= rsget("cardOptionName")

                FItemList(i).FcpnRegUserID      = rsget("cpnRegUserID")
	            FItemList(i).FcouponNo          = rsget("couponNo")
				i=i+1
				rsget.movenext
			loop
		End If
		rsget.close

	End Sub


	public Sub getCSGiftcardOrderDetail()
		dim sqlStr, i

		sqlStr = "EXEC [db_order].[dbo].[usp_cs_GiftCard_orderDetail] '" & FRectGiftOrderSerial & "' "
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        FResultcount = rsget.Recordcount

        if Not rsget.Eof then
			set FOneItem = new cGiftcardOrderItem

			FOneItem.FgiftOrderSerial	= rsget("giftOrderSerial")
			FOneItem.Fidx				= rsget("idx")
			FOneItem.FcardItemid		= rsget("cardItemid")
			FOneItem.FcardOption		= rsget("cardOption")
			FOneItem.FmasterCardCode	= rsget("masterCardCode")
			FOneItem.FmasterCheckCode	= rsget("masterCheckCode")
			FOneItem.Fuserid			= rsget("userid")
			FOneItem.Fbuyname			= rsget("buyname")
			FOneItem.Fbuyemail			= rsget("buyemail")
			FOneItem.Fbuyhp				= rsget("buyhp")
			FOneItem.FbuyPhone			= rsget("buyPhone")
			FOneItem.Fsendhp			= rsget("sendhp")
			FOneItem.Fsendemail			= rsget("sendemail")
			FOneItem.Freqhp				= rsget("reqhp")
			FOneItem.Freqemail			= rsget("reqemail")
			FOneItem.Ftotalsum			= rsget("totalsum")
			FOneItem.Fjumundiv			= rsget("jumundiv")
			FOneItem.Faccountdiv		= rsget("accountdiv")
			FOneItem.Fipkumname			= rsget("ipkumname")
			FOneItem.Faccountno			= rsget("accountno")
			FOneItem.Fipkumdiv			= rsget("ipkumdiv")
			FOneItem.Fipkumdate			= rsget("ipkumdate")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Fcancelyn			= rsget("cancelyn")
			FOneItem.Fpaydateid			= rsget("paydateid")
			FOneItem.Fresultmsg			= rsget("resultmsg")
			FOneItem.Fauthcode			= rsget("authcode")
			FOneItem.Fdiscountrate		= rsget("discountrate")
			FOneItem.Fsubtotalprice		= rsget("subtotalprice")
			FOneItem.Fmiletotalprice	= rsget("miletotalprice")
			FOneItem.Ftencardspend		= rsget("tencardspend")
			FOneItem.Fcashreceiptreq	= rsget("cashreceiptreq")
			FOneItem.FinsureCd			= rsget("insureCd")
			FOneItem.FinsureMsg			= rsget("insureMsg")
			FOneItem.Freferip			= rsget("referip")
			FOneItem.Fuserlevel			= rsget("userlevel")
			FOneItem.Fcanceldate		= rsget("canceldate")
			FOneItem.FsumPaymentEtc		= rsget("sumPaymentEtc")
			FOneItem.FsendDiv			= rsget("sendDiv")
			FOneItem.FbookingYn			= rsget("bookingYn")
			FOneItem.FbookingDate		= rsget("bookingDate")
			FOneItem.FsendDate			= rsget("sendDate")
			FOneItem.FdesignId			= rsget("designId")
			FOneItem.FMMSTitle			= rsget("MMSTitle")
			FOneItem.FMMSContent		= rsget("MMSContent")
			FOneItem.FemailTitle		= rsget("emailTitle")
			FOneItem.FemailContent		= rsget("emailContent")

			FOneItem.FSmallimage		= FGiftCardWebImgUrl & "/giftcard/small/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("smallImage")
			FOneItem.FCarditemname		= rsget("Carditemname")
			FOneItem.FCardinfo			= rsget("Cardinfo")
			FOneItem.FCarddesc			= rsget("Carddesc")
			FOneItem.FcardPrice			= rsget("cardPrice")
			FOneItem.FcardregDate		= rsget("cardregDate")
			FOneItem.Freguserid			= rsget("reguserid")
			FOneItem.FcardStatus		= rsget("cardStatus")
			FOneItem.FcardOptionName	= rsget("cardOptionName")

		End If
		rsget.close
	End Sub


	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		FGiftCardWebImgUrl = "http://webimage.10x10.co.kr"
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
