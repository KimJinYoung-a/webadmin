<%
'###########################################################
' Description : 주문 클래스
' Hieditor : 2009.04.17 이상구 생성
'			 2016.07.19 한용민 수정
'###########################################################

class CDesignerJumunList
	public Forderserial
	public Fjumundiv
	public Fuserid
	public Faccountdiv
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbuyname
	public Freqname
	public Freqphone
	public Freqhp
	public Fdeliverno
  	public Fdeliverytype
	public Fsitename
	public Fdiscountrate
	public FItemNo
	public Fitemcost
	public FCancelyn
	public FCurrPage
	Public FCurrState
	public FItemID
	public FItemName
	public FItemOption
	public FItemOptionStr
	public FBuycash
	public FMakerid
	public FUpcheBaesongDate
	public FIsUpcheBeasong
	public Fsongjangdiv
	public Fsongjangno
	public Fsongjangdivname
	public Ffindurl
	public fupchemanagecode

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#444400"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FFFF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
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
			IpkumDivName="주문접수"
		elseif Fipkumdiv="4" then
			IpkumDivName="결제완료"
		elseif Fipkumdiv="5" then
			IpkumDivName="배송통보"
		elseif Fipkumdiv="6" then
			IpkumDivName="배송준비"
		elseif Fipkumdiv="7" then
			IpkumDivName="일부출고"
		elseif Fipkumdiv="8" then
			IpkumDivName="상품출고"
		end if
	end Function

	Public function NormalUpcheDeliverState()
		 if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        NormalUpcheDeliverState = "주문접수"
		    else
			    NormalUpcheDeliverState = "결제완료"
			end if
		 elseif FCurrState="2" then
			 NormalUpcheDeliverState = "주문통보"
		 elseif FCurrState="3" then
			 NormalUpcheDeliverState = "주문확인"
		 elseif FCurrState="7" then
			 NormalUpcheDeliverState = "출고완료"
		 else
			 NormalUpcheDeliverState = ""
		 end if
	 end Function

	public function UpCheDeliverStateColor()
		if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        UpCheDeliverStateColor ="#444444"
		    else
			    UpCheDeliverStateColor ="#3300CC"
			end if

		elseif FCurrState="2" then
			UpCheDeliverStateColor="#336600"
		elseif FCurrState="3" then
			UpCheDeliverStateColor="#CC9933"
		elseif FCurrState="7" then
			UpCheDeliverStateColor="#FF0000"
		else
			UpCheDeliverStateColor="#000000"
		end if
	end function

    '' 플라워 지정일 시각
    public function GetReqTimeText()
        if IsNULL(Freqtime) then Exit function
        GetReqTimeText = Freqtime & "~" & (Freqtime+2) & "시 경"
    end function
end class

class CJumunDetailItem
	public Forderserial
	public Fdetailidx
	public FMakerid
	public Fitemid
	public Fitemoption
	public Fitemno
	public Fitemcost
    public FitemcostCouponNotApplied
    public FreducedPrice
	public Fbuycash
	public Fitemvat
	public Fmileage
	public Fcancelyn
	public FItemName
	public FItemoptionName
	public FImageList
	public FImageSmall
    public Fcurrstate
    public Fsongjangdiv
    public Fsongjangno
    public Fupcheconfirmdate
    public Fbeasongdate
    public Fisupchebeasong
    public Fissailitem
    public Foitemdiv
	public Frequiredetail
	public FcurrSellcash
	public FcurrBuycash
    public FOmwDiv
    public FoDlvType
	public Fvatinclude

	Public Fsongjangdivname
	Public Ffindurl

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
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CJumunDetail
	public FJumunDetailList()
	public FDetailCount

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
	end function

	public function BeasongPay()
		dim i, retVal : retVal=0
		for i=0 to FDetailCount-1
			if FJumunDetailList(i).FItemID=0 then
			    if isNULL(FJumunDetailList(i).Fitemcost) then
			        retVal = retVal + 0
			    else
    				retVal = retVal + CLNG(FJumunDetailList(i).Fitemcost*FJumunDetailList(i).Fitemno)
    			end if
			end if
		next
		BeasongPay = retVal
	end Function

	public function BeasongOptionStr()
		dim i
		for i=0 to FDetailCount-1
			if FJumunDetailList(i).FItemID=0 then
				BeasongOptionStr = BeasongCD2Name(FJumunDetailList(i).Fitemoption)
				Exit For
			end if
		next
	end function

	public sub SetDetailCount(byval v)
		FDetailCount = v
		redim preserve FJumunDetailList(v)
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CJumunMasterItem
	public ftemp_idx
	public fcustomnumber
	public Forderserial
	public Fjumundiv
	public Fuserid
	public Fuserlevel
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalvat
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public FCvbeadaldate
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
	public Fcomment
	public Fdeliverno
	public Fsitename
	public Fpaygatetid
	public Fdiscountrate
	public Fsubtotalprice
	public Fsubtotal
	public FAvgTotal
	public Fresultmsg
	public Frduserid
	public Fmiletotalprice
	public Fjungsanflag
	public Freqzipaddr
	public Fauthcode
	public Fcouponpay
	public freqemail
	public FDtlItemName
	public FDtlItemNo
	public FDtlItemOption
	public FDtlItemOptionName
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Freqdate
	public Freqtime
	public fgiftOrderSerial
    public FDlvcountryCode
    public FsumpaymentEtc       '''''2011-04 추가
    public FdlvPaySum           ''''2013/09/24

    public FsubtotalpriceCouponNotApplied
    public FallAtDiscountPrice
    public FRdsite
    public fpggubun
    public fmid
    public ftid
    public fauth_no
    public fp_rmesg1
    public fispay
    public fpayresultcode
    public fissuccess
    public fprice

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

    public function getMayItemCouponDiscount()
        getMayItemCouponDiscount = FsubtotalpriceCouponNotApplied-Fcouponpay-Fmiletotalprice-Fsubtotalprice-Fallatdiscountprice
    end function

    public function TotalMajorPaymentPrice()
        TotalMajorPaymentPrice = CLNG(FsubtotalPrice)-CLNG(FsumPaymentEtc)
    end function

    ''해외배송인지여부
	public function IsForeignDeliver()
        IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR")  and (FDlvcountryCode<>"ZZ")
    end function

	'/사용금지	'/공용 펑션에 공용 함수 사용할것.	'/2016.07.19 한용민
	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44DD44"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444FF"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#FF1111"   ''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#7D2448"   ''VIP GOLD
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

	'/사용금지	'/공용 펑션에 공용 함수 사용할것.	'/2016.07.19 한용민
	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Green"   		''Green
		elseif Fuserlevel="2" then
			GetUserLevelName = "Blue"   		''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelName = "VIP Silver"   	''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelName = "VIP Gold"   	''VIP GOLD
		elseif Fuserlevel="9" then
			GetUserLevelName = "Mania"  		'' mania
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

	public function GetRegDate()
		GetRegDate = FRegDate

		''CStr(year(FRegDate)) + "-" + CStr(Month(FRegDate)) + "-" + CStr(Day(FRegDate)) + " " + CStR(Hour(FRegDate)) + ":" + CStR(Min(FRegDate)) + ":" + CStR(second(FRegDate))
	end function

	public function UserIDName()
		if IsNull(fUserID) or (FUserID="") then
			UserIDName = "&nbsp;"
		else
			UserIDName = FUserID
		end if
	end function

	public function IsAvailAndIpkumOK()
		IsAvailAndIpkumOK = (CInt(Fipkumdiv)>3) and IsAvailJumun
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	public function IpkumDivColor()
		if FjumunDiv="9" then
			IpkumDivColor = "#FF0000"
		else
			if Fipkumdiv="0" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="1" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="2" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="3" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="4" then
				IpkumDivColor="#0000FF"
			elseif Fipkumdiv="5" then
				IpkumDivColor="#444400"
			elseif Fipkumdiv="6" then
				IpkumDivColor="#FFFF00"
			elseif Fipkumdiv="7" then
				IpkumDivColor="#004444"
			elseif Fipkumdiv="8" then
				IpkumDivColor="#FF00FF"
			end if
		end if
	end function

	public function SiteNameColor()
		if Fsitename="uto" then
			SiteNameColor = "#55AA22"
		elseif Fsitename="cara" then
			SiteNameColor = "#225555"
		elseif Fsitename="emoden" then
			SiteNameColor = "#992255"
		elseif Fsitename="netian" then
			SiteNameColor = "#AA22AA"
		elseif Fsitename="miclub" then
			SiteNameColor = "#22AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function

	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		''elseif FSubtotalPrice>50000 then
		''	SubTotalColor = "#33AAAA"
		else
			SubTotalColor = "#000000"
		end if
	end function

	public function CancelYnColor()
		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function

	public function CancelYnName()
		if FCancelYn="D" then
			CancelYnName = "삭제"
		elseif UCase(FCancelYn)="Y" then
			CancelYnName = "취소"
		elseif FCancelYn="N" then
			CancelYnName = "정상"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="외부몰"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+신용"
		elseif Faccountdiv="400" then
			JumunMethodName="핸드폰결제"
		end if
	end function

	Public function IpkumDivName()
		if FjumunDiv="9" then
			IpkumDivName = "마이너스"
		else
			if Fipkumdiv="0" then
				IpkumDivName="주문대기"
			elseif Fipkumdiv="1" then
				IpkumDivName="주문실패"
			elseif Fipkumdiv="2" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="3" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="4" then
				IpkumDivName="결제완료"
			elseif Fipkumdiv="5" then
				IpkumDivName="배송통보"
			elseif Fipkumdiv="6" then
				IpkumDivName="배송준비"
			elseif Fipkumdiv="7" then
				IpkumDivName="일부출고"
			elseif Fipkumdiv="8" then
				IpkumDivName="출고완료"
			end if
		end if
	end function

    '' 플라워 지정일 시각
    public function GetReqTimeText()
        if IsNULL(Freqtime) then Exit function
        GetReqTimeText = Freqtime & "~" & (Freqtime+2) & "시 경"
    end function

end Class

class CJumunMaster
	public FMasterItemList()
	public FMasterItemList2()
	public FJumunDetail
	public FSubtotal
	public FTotalCount
	public FTotalBuyCash
	public FAvgTotal
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectSearchtype
	public FRectSearchtype01
	public FRectSearchtype02
	public FRectOrderSerial
	public FRectgiftOrderSerial
	public FRectUserID
	public FRectckdate
	public FRectBuyname
	public FRectReqName
	public FRectIpkumName
	public FRectSubTotalPrice
	public FRectAuthCode
	public FRectRegStart
	public FRectRegEnd
	public FRectDelNoSearch
	public FRectIpkumDiv2
	public FRectIpkumDiv4
	public FRectSiteName
	public FRectRdSite
	public FRectOnlyIpkumDiv
    public FRectckpointsearch
	public FRectNotThisSite
	public FRectNoViewPoint
	public FRectDesignerID
	public FRectItemid
	public FRectItemName
	public FRectOnlyOutMall
	public FRectDateType
	public FRectOrderBy
	public FRectDispY
	public FRectSellY
	public FRectOnlyPoint
	public FRectIpkumOrJumun
	public FRectBeasongNotFinish
	public Fnotitemlist
	public Fitemlist
	public FRectIpkumDiv4before
	public FRectOldJumun
	public FRectcancelyn
    public FRectIpkumDiv
    public FRectMinusOrderInclude
	public FRectreqdate
    public FRectIsUpcheBeasong
	public FRectreqHp
    public FRectIsFlower
    public FRectIsMinus
    public FRectIsForeign
    public FRectIsMilitary
	public FRectResultMessage
	public FRectIsPay
	public FRectIsSuccess
	public FRectIsSendGift

    ''2012/05추가
    public FRectJumunDiv
    public FRectBuyHp
    public FRectBuyEmail

    ''2012/09추가
    public FRectCDL
    public FRectCDM
    public FRectCDS
    public FRectBrandPurchaseType
    public FIsMDPick
    public FIsRdSite
    public FIsMDPickMo

    ''2017/11/22
    public FIsMDPickMoArr

    '2013/09/24
    public FRectQryDLVsum
    public FRectChannelDiv
	public FRectDispCate
    public FRectSellChannelDiv

    '2015/12/21
    public FRectPojangOrder

    '2017/03/29
    public FRectExcept3pl

	public FRectBankOnly
	public FRectpaygateid
	public FRectprice
	public FRecttemp_idx

	Private Sub Class_Initialize()
'		redim preserve FMasterItemList(0)
		redim FMasterItemList(0)
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FRectBankOnly = "Y"
	End Sub

	Private Sub Class_Terminate()
	End Sub
	public function Null2Zoro(byval v)
		if IsNull(v) then
			Null2Zoro = 0
		else
			Null2Zoro = v
		end if
	end function

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FJumunDetail.FJumunDetailList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FJumunDetail.FJumunDetailList(i).FItemID)
	end function

	public function IsAllTenBeasong()
		dim sqlStr
		sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " where makerid='" + session("ssBctID") + "'"
		sqlStr = sqlStr + " and deliverytype<>'1'"
		sqlStr = sqlStr + " and deliverytype<>'4'"
		sqlStr = sqlStr + " and sellyn<>'N'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IsAllTenBeasong = (rsget("cnt")<1)
		rsget.close
	end function

	public sub SearchTargetItemJumunList()
		dim sqlStr
		dim i

		sqlStr = "select top 1000 m.orderserial, m.buyname, m.reqname, m.ipkumdiv, m.jumundiv,"
		sqlStr = sqlStr + " convert(varchar(10),m.regdate,21) as regdate, convert(varchar(10),m.ipkumdate,21) as ipkumdate, m.comment , d.itemname, d.itemno,"
		sqlStr = sqlStr + " d.itemoption, d.itemoptionname"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"

		if FRectIpkumOrJumun="j" then
			sqlStr = sqlStr + " and m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.regdate<'" + FRectRegEnd + "'"
		else
			sqlStr = sqlStr + " and m.ipkumdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.ipkumdate<'" + FRectRegEnd + "'"
		end if

		if FRectBeasongNotFinish="on" then
			sqlStr = sqlStr + " and m.ipkumdiv>4"
			sqlStr = sqlStr + " and m.ipkumdiv<8"
		else
			sqlStr = sqlStr + " and m.ipkumdiv>4"
		end if

		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		if FRectItemid<>"" then
		sqlStr = sqlStr + " and d.itemid=" + FRectItemid + ""
		end if
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " order by m.idx"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fregdate		= rsget("regdate")
			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")

			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
			FMasterItemList(i).FDtlItemName = db2Html(rsget("itemname"))
			FMasterItemList(i).FDtlItemNo	= rsget("itemno")
			FMasterItemList(i).FDtlItemOption = rsget("itemoption")
			FMasterItemList(i).FDtlItemOptionName = db2Html(rsget("itemoptionname"))
			rsget.movenext
			i=i+1
		loop
		rsget.Close

	end Sub

	public sub SearchOneJumunDetail(byval idx)
		dim sqlStr, i

		sqlStr = "select d.*, convert(varchar(19),d.upcheconfirmdate,21) as cvupcheconfirmdate, convert(varchar(19),d.beasongdate,21) as cvbeasongdate, i.smallimage,i.listimage, i.sellcash as currsellcash, i.buycash as currbuycash"
		if (FRectOldJumun="on") then
		sqlStr = sqlStr + "  from [db_log].[dbo].tbl_old_order_detail_2003 d"
		else
		sqlStr = sqlStr + "  from [db_order].[dbo].tbl_order_detail d"
	    end if
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where d.idx=" + CStr(idx)

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		set FJumunDetail = new CJumunDetailItem

		if Not rsget.Eof then
			FJumunDetail.Forderserial = rsget("orderserial")
			FJumunDetail.Fdetailidx			= rsget("idx")
			FJumunDetail.Fmakerid      = rsget("makerid")
			FJumunDetail.Fitemid      = rsget("itemid")
			FJumunDetail.Fitemoption  = rsget("itemoption")
			FJumunDetail.Fitemno      = rsget("itemno")
			FJumunDetail.Fitemcost    = rsget("itemcost")
            FJumunDetail.FitemcostCouponNotApplied    = rsget("itemcostCouponNotApplied")
            FJumunDetail.FreducedPrice    = rsget("reducedPrice")
			FJumunDetail.Fbuycash     = rsget("buycash")
			FJumunDetail.Fmileage     = rsget("mileage")
			'FJumunDetail.Fcosttotal   = rsget("costtotal")
			'FJumunDetail.Forderdate   = rsget("orderdate")
			FJumunDetail.Fcancelyn    = rsget("cancelyn")

			FJumunDetail.FItemName    = db2html(rsget("itemname"))
			FJumunDetail.FImageList    = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FJumunDetail.Fitemid) + "/" + rsget("listimage")
			FJumunDetail.FImageSmall    = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FJumunDetail.Fitemid) + "/" + rsget("smallimage")

			FJumunDetail.FItemoptionName = db2html(rsget("itemoptionname"))

			FJumunDetail.Fcurrstate     = rsget("currstate")
			FJumunDetail.Fsongjangdiv   = rsget("songjangdiv")
			FJumunDetail.Fsongjangno    = rsget("songjangno")
			FJumunDetail.Fupcheconfirmdate = rsget("cvupcheconfirmdate")
			FJumunDetail.Fbeasongdate   = rsget("cvbeasongdate")
			FJumunDetail.Fisupchebeasong= rsget("isupchebeasong")
			FJumunDetail.Fissailitem    = rsget("issailitem")

			FJumunDetail.Frequiredetail = db2html(rsget("requiredetail"))

			FJumunDetail.FcurrSellcash	= rsget("currsellcash")
			FJumunDetail.FcurrBuycash	= rsget("currbuycash")
            FJumunDetail.Foitemdiv      = rsget("oitemdiv")

            FJumunDetail.FOmwDiv        = rsget("omwdiv")
            FJumunDetail.FODlvType      = rsget("odlvtype")

			FJumunDetail.Fvatinclude      = rsget("vatinclude")
		end if

		rsget.close

	end sub

	public sub SearchJumunDetail(byval orderserial)
		dim sqlStr
		dim i
		dim tblMasterName, tblDetailName

		tblMasterName = "[db_order].[dbo].tbl_order_master"
		tblDetailName = "[db_order].[dbo].tbl_order_detail"
		if (FRectOldJumun = "on") then
			tblMasterName = "[db_log].[dbo].tbl_old_order_master_2003"
			tblDetailName = "[db_log].[dbo].tbl_old_order_detail_2003"
		end if

		sqlStr = "select d.idx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.itemvat,d.mileage,d.cancelyn,d.buycash,"
		sqlStr = sqlStr + " d.itemname, d.makerid, i.listimage as imglist, i.smallimage as imgsmall, d.itemoptionname as codeview"
		sqlStr = sqlStr + " , d.currstate, d.songjangdiv, d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem"
		sqlStr = sqlStr + " , isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail"
		sqlStr = sqlStr + " ,s.divname as songjangdivname, s.findurl"
		sqlStr = sqlStr + " from " & CStr(tblDetailName) & " d "
		sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i on d.itemid=i.itemid "
		sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_songjang_div s on d.songjangdiv=s.divcd"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON d.idx = dd.detailidx" & vbcrlf
		sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		set FJumunDetail = new CJumunDetail
		FJumunDetail.SetDetailCount rsget.RecordCount

		i=0
		do until rsget.eof
			set FJumunDetail.FJumunDetailList(i) = new CJumunDetailItem

			FJumunDetail.FJumunDetailList(i).Forderserial = CStr(orderserial)
			FJumunDetail.FJumunDetailList(i).Fdetailidx			= rsget("idx")
			FJumunDetail.FJumunDetailList(i).Fmakerid      = rsget("makerid")
			FJumunDetail.FJumunDetailList(i).Fitemid      = rsget("itemid")
			FJumunDetail.FJumunDetailList(i).Fitemoption  = rsget("itemoption")
			FJumunDetail.FJumunDetailList(i).Fitemno      = rsget("itemno")
			FJumunDetail.FJumunDetailList(i).Fitemcost    = rsget("itemcost")
			FJumunDetail.FJumunDetailList(i).Fbuycash    = rsget("buycash")
			FJumunDetail.FJumunDetailList(i).Fitemvat     = rsget("itemvat")
			FJumunDetail.FJumunDetailList(i).Fmileage     = rsget("mileage")
			''FJumunDetail.FJumunDetailList(i).Fcosttotal   = rsget("costtotal")
			''FJumunDetail.FJumunDetailList(i).Forderdate   = rsget("orderdate")
			FJumunDetail.FJumunDetailList(i).Fcancelyn    = rsget("cancelyn")

			FJumunDetail.FJumunDetailList(i).FItemName    = db2html(rsget("itemname"))
			FJumunDetail.FJumunDetailList(i).FImageList    = webImgUrl + "/image/list/" + GetImageFolerName(i) + "/" + rsget("imglist")
			FJumunDetail.FJumunDetailList(i).FImageSmall    = webImgUrl + "/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")

			if IsNull(rsget("codeview")) then
				FJumunDetail.FJumunDetailList(i).FItemoptionName = "-"
			else
				FJumunDetail.FJumunDetailList(i).FItemoptionName = db2html(rsget("codeview"))
			end if

			FJumunDetail.FJumunDetailList(i).Fcurrstate     = rsget("currstate")
			FJumunDetail.FJumunDetailList(i).Fsongjangdiv   = rsget("songjangdiv")
			FJumunDetail.FJumunDetailList(i).Fsongjangno    = rsget("songjangno")
			FJumunDetail.FJumunDetailList(i).Fbeasongdate   = rsget("beasongdate")
			FJumunDetail.FJumunDetailList(i).Fisupchebeasong= rsget("isupchebeasong")
			FJumunDetail.FJumunDetailList(i).Fissailitem    = rsget("issailitem")

			FJumunDetail.FJumunDetailList(i).Frequiredetail = db2html(rsget("requiredetail"))

			FJumunDetail.FJumunDetailList(i).Fsongjangdivname  = db2html(rsget("songjangdivname"))
            FJumunDetail.FJumunDetailList(i).Ffindurl          = db2html(rsget("findurl"))

			rsget.movenext
			i=i+1
		loop
		rsget.close
	end sub

	public sub SearchNotDeliverList()
		dim sqlStr,wheredetail
		dim i
		wheredetail = ""

		if (FRectOnlyIpkumDiv<>"") then
			wheredetail = " and m.regdate>'2005-11-01'"
			wheredetail = wheredetail + " and m.ipkumdiv=" + CStr(FRectOnlyIpkumDiv)
		else
			wheredetail = " and m.regdate>'2005-11-01'"
			wheredetail = wheredetail + " and m.ipkumdiv>3"
		end if

		if (FRectNotThisSite<>"") then
			wheredetail = wheredetail + " and m.sitename<>'" + FRectNotThisSite + "'"
		end if

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if
		wheredetail = wheredetail + " and m.cancelyn='N'"

		''########## 총 갯수 ################''
		sqlStr = "select count(m.orderserial) as cnt"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where orderserial<>''"
		sqlStr = sqlStr + wheredetail

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		''########## 데이타 ################''
		sqlStr = "select top " + CStr(FPageSize) + " m.orderserial, m.buyname, m.userid, m.accountdiv, m.sitename, "
		sqlStr = sqlStr + " m.totalsum, m.subtotalprice, m.ipkumdiv, m.regdate, "
		sqlStr = sqlStr + " m.discountrate, m.buyname, m.reqname, m.cancelyn"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where orderserial not in ("
		sqlStr = sqlStr + " select top " + CStr((FCurrPage-1)*FPageSize)  + " orderserial from [db_order].[dbo].tbl_order_master "
		sqlStr = sqlStr + " where orderserial<>''"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by regdate desc"
		sqlStr = sqlStr + ")"

		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by regdate desc"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsget("regdate")
			FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	public sub SearchDirectBuyList()
		dim sqlStr,wheredetail
		dim i
		wheredetail = ""

		if (FRectOnlyIpkumDiv<>"") then
			wheredetail = " where m.regdate>'2005-11-01'"
			wheredetail = wheredetail + " and m.ipkumdiv=" + CStr(FRectOnlyIpkumDiv)
		else
			wheredetail = " where m.regdate>'2005-11-01'"
			wheredetail = wheredetail + " and m.ipkumdiv>3"
		end if

		wheredetail = wheredetail + " and m.orderserial=d.orderserial"
		wheredetail = wheredetail + " and d.itemid=0"
		wheredetail = wheredetail + " and d.itemoption='0301'"
		wheredetail = wheredetail + " and m.cancelyn='N'"
		wheredetail = wheredetail + " and d.cancelyn<>'Y'"

		sqlStr = "select count(m.orderserial) as cnt"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"

		sqlStr = sqlStr + wheredetail

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top 100 m.orderserial, m.buyname, m.userid, m.accountdiv, m.sitename, "
		sqlStr = sqlStr + " m.totalsum, m.subtotalprice, m.ipkumdiv, m.regdate, "
		sqlStr = sqlStr + " m.discountrate, m.buyname, m.reqname, m.cancelyn"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"

		sqlStr = sqlStr + wheredetail

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsget("regdate")
			FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	public Sub SearchJumunListByDesigner()
		dim sqlStr
		dim i
		dim wheredetail

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectUserID<>"") then
			wheredetail = wheredetail + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			wheredetail = wheredetail + " and m.buyname = '" + FRectBuyname + "'"
		end if

		if (FRectReqName<>"") then
			wheredetail = wheredetail + " and m.reqname = '" + FRectReqName + "'"
		end if

		if (FRectSubTotalPrice<>"") then
			wheredetail = wheredetail + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectRegStart<>"") then
			if FRectDateType="ipkumil" then
				wheredetail = wheredetail + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			elseif FRectDateType="upbeasongdate" then
				'기존것,좀이상함;;20110804강준구수정wheredetail = wheredetail + " and ((d.isupchebeasong='Y') and (d.beasongdate >='" + CStr(FRectRegStart) + "')) "
				wheredetail = wheredetail + " and d.beasongdate >='" + CStr(FRectRegStart) + "' "
			elseif FRectDateType="tenbeasongdate" then
				'기존것,좀이상함;;20110804강준구수정wheredetail = wheredetail + " and ((d.isupchebeasong='N') and (m.beadaldate >='" + CStr(FRectRegStart) + "'))"
				wheredetail = wheredetail + " and d.beasongdate >='" + CStr(FRectRegStart) + "'"
			else
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if FRectDateType="ipkumil" then
				wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif FRectDateType="upbeasongdate" then
				'기존것,좀이상함;;20110804강준구수정wheredetail = wheredetail + " and ((d.isupchebeasong='Y') and (d.beasongdate <'" + CStr(FRectRegEnd) + "')) "
				wheredetail = wheredetail + " and d.beasongdate <'" + CStr(FRectRegEnd) + "' "
			elseif FRectDateType="tenbeasongdate" then
				'기존것,좀이상함;;20110804강준구수정wheredetail = wheredetail + " and ((d.isupchebeasong='N') and (m.beadaldate <'" + CStr(FRectRegEnd) + "'))"
				wheredetail = wheredetail + " and d.beasongdate <'" + CStr(FRectRegEnd) + "'"
			else
				wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectIpkumName<>"") then
			wheredetail = wheredetail + " and m.accountname = '" + FRectIpkumName + "'"
		end if

		if (FRectSiteName<>"") then
			wheredetail = wheredetail + " and m.sitename ='" + FRectSiteName + "'"
		end if

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if

		if (FRectItemID<>"") then
			wheredetail = wheredetail + " and d.itemid="&FRectItemID&""
		end if

        if (FRectIsUpcheBeasong<>"") then
            wheredetail = wheredetail + " and d.isupchebeasong='"&FRectIsUpcheBeasong&"'"
		end if

        if (FRectIpkumdiv<>"") then
            wheredetail = wheredetail + " and m.ipkumdiv='"&FRectIpkumdiv&"'"
		end if

		sqlStr = "select count(m.orderserial) as cnt, sum(d.buycash*d.itemno) as totalbuycash"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock) "
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d with (nolock) "
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1'"             ''2009 변경, 주문접수건도 표시
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalBuyCash = rsget("totalbuycash")
		    FTotalCount = rsget("cnt")

		    if IsNULL(FTotalBuyCash) then FTotalBuyCash=0
		rsget.Close

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " d.orderserial, m.buyname,m.reqname, m.jumundiv, m.userid,"
		sqlStr = sqlStr + " m.ipkumdiv, m.ipkumdate, m.accountdiv, m.regdate, m.reqphone, m.reqhp, m.deliverno, "
		sqlStr = sqlStr + " m.sitename, m.discountrate, m.cancelyn, m.authcode "
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoption, d.itemno, d.itemoptionname as optname, d.itemcost,"
		sqlStr = sqlStr + " d.beasongdate,d.isupchebeasong, d.buycash, d.currstate"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock) , "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d with (nolock) "
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1'"             ''2009 변경, 주문접수건도 표시
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by d.idx desc"
'response.write sqlStr
		rsget.pagesize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
    		do until rsget.eof
    			set FMasterItemList(i) = new CDesignerJumunList
    			FMasterItemList(i).Forderserial = rsget("orderserial")
    			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
    			FMasterItemList(i).Fuserid		= rsget("userid")
    			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
    			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
    			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
    			FMasterItemList(i).Fregdate		= rsget("regdate")
    			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
    			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
    			FMasterItemList(i).Freqphone	= rsget("reqphone")
    			FMasterItemList(i).Freqhp		= rsget("reqhp")
    			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
    			FMasterItemList(i).Fsitename	= rsget("sitename")
    			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
    			FMasterItemList(i).FCancelyn	= rsget("cancelyn")
    			FMasterItemList(i).Fauthcode		= rsget("authcode")

    			FMasterItemList(i).FItemID       = rsget("itemid")
    			FMasterItemList(i).FItemName     = db2Html(rsget("itemname"))
    			FMasterItemList(i).FItemOption   = rsget("itemoption")
    			FMasterItemList(i).FItemOptionStr= db2Html(rsget("optname"))
    			FMasterItemList(i).FItemNo     = rsget("itemno")
    			FMasterItemList(i).Fitemcost     = rsget("itemcost")

    			FMasterItemList(i).FUpcheBaesongDate     = rsget("beasongdate")
    			FMasterItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
    			FMasterItemList(i).Fbuycash = rsget("buycash")
    			FMasterItemList(i).FCurrState		 = rsget("currstate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//designer/jumunmaster/jumunlist.asp	'/2016.06.02 한용민 수정(페이징 방식 변경. 일정시간에 기계가 쿼리해감.부하가 심함)
	public Sub SearchJumunListByDesignerNew()
		dim sqlStr, i

		if (FRectOldJumun = "on") and (FRectOrderSerial = "") then
			'// 해킹대비
			response.write "잘못된 접근입니다. - 과거내역을 검색하려면 주문번호를 입력해야 합니다."
			response.end
		end if

		if (FRectItemID = "") then
			FRectItemID = 0
		end if

		if (FRectSubTotalPrice = "") then
			FRectSubTotalPrice = 0
		end if

		sqlStr = " exec [db_order].[dbo].[usp_Ten_GetOrderCountByDesigner] '" + CStr(FRectDesignerID) + "', '" + CStr(FRectOldJumun) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectOrderSerial) + "', '" + CStr(FRectUserID) + "', '" + CStr(FRectBuyname) + "', '" + CStr(FRectReqName) + "', '" + CStr(FRectIpkumName) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectSiteName) + "', '" + CStr(FRectNoViewPoint) + "', '" + CStr(FRectIsUpcheBeasong) + "', " + CStr(FRectItemID) + ", " + CStr(FRectSubTotalPrice) + ", '" + CStr(FRectDateType) + "', '" + CStr(FRectRegStart) + "', '" + CStr(FRectRegEnd) + "' "

		'response.write sqlStr & "<Br>"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalBuyCash = rsget("totalbuycash")
		    FTotalCount = rsget("cnt")

		    if IsNULL(FTotalBuyCash) then FTotalBuyCash=0
		rsget.Close

		sqlStr = " exec [db_order].[dbo].[usp_Ten_GetOrderListByDesigner] '"&CStr((FPageSize*(FCurrPage-1)) + 1)&"','"&CStr(FPageSize*FCurrPage)&"', '" + CStr(FRectDesignerID) + "', '" + CStr(FRectOldJumun) + "' "
		'sqlStr = " exec [db_order].[dbo].[usp_Ten_GetOrderListByDesigner] " + CStr(FPageSize * FCurrPage) + ", '" + CStr(FRectDesignerID) + "', '" + CStr(FRectOldJumun) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectOrderSerial) + "', '" + CStr(FRectUserID) + "', '" + CStr(FRectBuyname) + "', '" + CStr(FRectReqName) + "', '" + CStr(FRectIpkumName) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectSiteName) + "', '" + CStr(FRectNoViewPoint) + "', '" + CStr(FRectIsUpcheBeasong) + "', " + CStr(FRectItemID) + ", " + CStr(FRectSubTotalPrice) + ", '" + CStr(FRectDateType) + "', '" + CStr(FRectRegStart) + "', '" + CStr(FRectRegEnd) + "' "

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0
'		FtotalPage =  CInt(FTotalCount\FPageSize)
'		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
'			FtotalPage = FtotalPage +1
'		end if
'		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
'        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF  then
			'rsget.absolutepage = FCurrPage
    		do until rsget.eof
    			set FMasterItemList(i) = new CDesignerJumunList
	    			FMasterItemList(i).Forderserial = rsget("orderserial")
	    			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
	    			FMasterItemList(i).Fuserid		= rsget("userid")
	    			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
	    			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
	    			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
	    			FMasterItemList(i).Fregdate		= rsget("regdate")
	    			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
	    			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
	    			FMasterItemList(i).Freqphone	= rsget("reqphone")
	    			FMasterItemList(i).Freqhp		= rsget("reqhp")
	    			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
	    			FMasterItemList(i).Fsitename	= rsget("sitename")
	    			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
	    			FMasterItemList(i).FCancelyn		= rsget("cancelyn")
	    			FMasterItemList(i).FItemID       	= rsget("itemid")
	    			FMasterItemList(i).FItemName     	= db2Html(rsget("itemname"))
	    			FMasterItemList(i).FItemOption   	= rsget("itemoption")
	    			FMasterItemList(i).FItemOptionStr	= db2Html(rsget("optname"))
	    			FMasterItemList(i).FItemNo     		= rsget("itemno")
	    			FMasterItemList(i).Fitemcost     	= rsget("itemcost")
	    			FMasterItemList(i).FUpcheBaesongDate    = rsget("beasongdate")
	    			FMasterItemList(i).FIsUpcheBeasong 		= rsget("isupchebeasong")
	    			FMasterItemList(i).Fbuycash 			= rsget("buycash")
	    			FMasterItemList(i).FCurrState		 	= rsget("currstate")

					FMasterItemList(i).Fsongjangdiv		 	= rsget("songjangdiv")
					FMasterItemList(i).Fsongjangno		 	= rsget("songjangno")
					FMasterItemList(i).Fsongjangdivname  	= db2html(rsget("songjangdivname"))
					FMasterItemList(i).Ffindurl          	= db2html(rsget("findurl"))
					FMasterItemList(i).fupchemanagecode		 	= rsget("upchemanagecode")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchOnlyOnJumunList()
		dim sqlStr,i, wheredetail

		wheredetail = ""

		wheredetail = wheredetail + " and m.regdate>'" + FRectRegStart + "'"
		wheredetail = wheredetail + " and m.cancelyn ='N'"
		wheredetail = wheredetail + " and m.ipkumdiv='4'"

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if

		if (FRectOnlyPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv='30'"
		end if

		sqlStr = "select count(T.orderserial) as cnt, sum(T.subtotalprice) as subtotal , avg(T.subtotalprice) as avgtotal "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " select m.orderserial, m.subtotalprice,  count(d.idx) as dcnt "
		sqlStr = sqlStr + "  from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + "  ,[db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by m.orderserial, m.subtotalprice"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where T.dcnt=1"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")

		FSubtotal = rsget("subtotal")
		FAvgTotal = rsget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize) + "T.* "
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " select  m.idx, m.orderserial, m.jumundiv, "
		sqlStr = sqlStr + " m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, "
		sqlStr = sqlStr + " count(d.idx) as dcnt "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m ,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " where m.orderserial=d.orderserial "
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid<>0 "
		sqlStr = sqlStr + " and d.cancelyn<>'Y' "
		sqlStr = sqlStr + " group by m.idx, m.orderserial, m.jumundiv, "
		sqlStr = sqlStr + " m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " convert(varchar,m.regdate,20)"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where T.dcnt=1"
		sqlStr = sqlStr + " order by T.idx"
'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsget("cvreg")
			FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	public Sub SearchQuickJumunList()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		wheredetail = wheredetail + " and m.regdate>'" + FRectRegStart + "'"
		wheredetail = wheredetail + " and m.cancelyn ='N'"
		wheredetail = wheredetail + " and m.ipkumdiv='4'"

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if

		if (FRectOnlyPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv='30'"
		end if

		sqlStr = "select count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_order].[dbo].tbl_order_master m WITH(READUNCOMMITTED)"
		sqlStr = sqlStr + " where m.orderserial in ("
		sqlStr = sqlStr + " select top 100 m.orderserial"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid=0"
		sqlStr = sqlStr + " and d.itemoption<>'0101'"
		sqlStr = sqlStr + " and d.itemoption<>'0501'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by m.orderserial desc"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + wheredetail
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")

		FSubtotal = rsget("subtotal")
		FAvgTotal = rsget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.orderserial in ("
		sqlStr = sqlStr + " select top 100 m.orderserial"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid=0"
		sqlStr = sqlStr + " and d.itemoption<>'0101'"
		sqlStr = sqlStr + " and d.itemoption<>'0501'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by m.orderserial desc"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by idx"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsget("cvreg")
			FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))
			FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsget("miletotalprice"))
			FMasterItemList(i).Fjungsanflag		= rsget("jungsanflag")
			FMasterItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
			FMasterItemList(i).Fauthcode		= rsget("authcode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	public Sub SearchBaljuJumunList()
		dim sqlStr
		dim i, wheredetail

		wheredetail = wheredetail + " and m.regdate>'" + FRectRegStart + "'"
		wheredetail = wheredetail + " and m.cancelyn ='N'"
		wheredetail = wheredetail + " and m.ipkumdiv='4'"

		'if (FRectNoViewPoint<>"") then
		'	wheredetail = wheredetail + " and m.accountdiv<>'30'"
		'end if

		if (FRectOnlyPoint<>"") then
			wheredetail = wheredetail + " and m.sitename in ('dnshop','auction')"
		end if

		if Fnotitemlist<>"" then
			sqlStr = "select count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_order].[dbo].tbl_order_master m WITH(READUNCOMMITTED)"
			sqlStr = sqlStr + " where orderserial not in ("
			sqlStr = sqlStr + " select distinct m.orderserial from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fnotitemlist + ")"
			sqlStr = sqlStr + " )"
			sqlStr = sqlStr + wheredetail
		elseif Fitemlist<>"" then
			sqlStr = "select count(distinct m.orderserial) as cnt, sum(distinct m.subtotalprice) as subtotal , avg(distinct m.subtotalprice) as avgtotal  from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fitemlist + ")"
		else
			sqlStr = "select count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_order].[dbo].tbl_order_master m WITH(READUNCOMMITTED)"
			sqlStr = sqlStr + " where m.idx<>0"
			sqlStr = sqlStr + wheredetail

		end if
'response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")

		FSubtotal = rsget("subtotal")
		FAvgTotal = rsget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsget.Close

		if Fnotitemlist<>"" then
			sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + " where orderserial not in ("
			sqlStr = sqlStr + " select distinct m.orderserial from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fnotitemlist + ")"
			sqlStr = sqlStr + " )"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " order by idx "
		elseif Fitemlist<>"" then
			sqlStr = "select distinct top " + CStr(FPageSize) + " "
			sqlStr = sqlStr + " m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv,"
			sqlStr = sqlStr + " m.accountno, m.totalvat, m.totalcost, m.totalmileage, m.totalsum,"
			sqlStr = sqlStr + " m.ipkumdiv,m.ipkumdate,m.beadaldiv,m.beadaldate,m.cancelyn,"
			sqlStr = sqlStr + " m.buyname,m.buyphone,m.buyhp,"
			sqlStr = sqlStr + " m.buyemail,m.reqname,m.reqzipcode,m.reqaddress,m.reqphone,"
			sqlStr = sqlStr + " m.reqhp,m.deliverno,m.sitename,m.paygatetid,"
			sqlStr = sqlStr + " m.discountrate,m.subtotalprice,m.resultmsg,m.rduserid,"
			sqlStr = sqlStr + " m.milelogid,m.miletotalprice,m.jungsanflag,m.reqzipaddr,m.authcode,"

			sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fitemlist + ")"
			sqlStr = sqlStr + " order by m.idx "
		else
			sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + " where m.idx<>0"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " order by idx "


		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Faccountno	= rsget("accountno")
			FMasterItemList(i).Ftotalvat	= Null2Zoro(rsget("totalvat"))
			FMasterItemList(i).Ftotalmileage= Null2Zoro(rsget("totalmileage"))
			FMasterItemList(i).Ftotalsum	= rsget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsget("cvreg")
			FMasterItemList(i).Fbeadaldiv	= rsget("beadaldiv")
			FMasterItemList(i).Fbeadaldate	= rsget("beadaldate")
			FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Fbuyphone	= rsget("buyphone")
			FMasterItemList(i).Fbuyhp		= rsget("buyhp")
			FMasterItemList(i).Fbuyemail	= rsget("buyemail")
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Freqzipcode	= rsget("reqzipcode")
			FMasterItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			FMasterItemList(i).Freqphone	= rsget("reqphone")
			FMasterItemList(i).Freqhp		= rsget("reqhp")
			''/FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fpaygatetid	= rsget("paygatetid")
			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))
			FMasterItemList(i).Fresultmsg		= rsget("resultmsg")
			FMasterItemList(i).Frduserid		= rsget("rduserid")
			FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsget("miletotalprice"))
			FMasterItemList(i).Fjungsanflag		= rsget("jungsanflag")
			FMasterItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
			FMasterItemList(i).Fauthcode		= rsget("authcode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	public Sub SearchJumunList()
		dim sqlStr, wheredetail, i, tblMasterName, tblDetailName

		tblMasterName = "[db_order].[dbo].tbl_order_master"
		tblDetailName = "[db_order].[dbo].tbl_order_detail"
		if (FRectOldJumun = "on") then
			tblMasterName = "[db_log].[dbo].tbl_old_order_master_2003"
			tblDetailName = "[db_log].[dbo].tbl_old_order_detail_2003"
		end if

		wheredetail = ""

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectUserID<>"") then
			wheredetail = wheredetail + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			wheredetail = wheredetail + " and m.buyname = '" + FRectBuyname + "'"
		end if

		if (FRectReqName<>"") then
			wheredetail = wheredetail + " and m.reqname = '" + FRectReqName + "'"
		end if

		if (FRectSubTotalPrice<>"") then
			wheredetail = wheredetail + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and m.cancelyn ='N'"
		end if

		if (FRectcancelyn<>"") then
			wheredetail = wheredetail + " and m.cancelyn ='"&FRectcancelyn&"'"
		end if

        if (FRectIpkumdiv<>"") then
            wheredetail = wheredetail + " and m.ipkumdiv='" & FRectIpkumdiv & "'"
        end if

		if (FRectIpkumDiv2<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv>='2'"
		end if

		if (FRectIpkumDiv4<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv>='4'"
		end if

		if (FRectIpkumDiv4before<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv < '4'"
		end if

		if (FRectOnlyIpkumDiv<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv=" + CStr(FRectOnlyIpkumDiv)
		end if

		if (FRectIpkumName<>"") then
			wheredetail = wheredetail + " and m.accountname = '" + FRectIpkumName + "'"
		end if

		if (FRectSiteName<>"") then
			If FRectSiteName = "shoplinker" Then
				wheredetail = wheredetail + " and m.sitename in (SELECT makerid FROM db_item.dbo.tbl_Shoplinker_OutmallControl WHERE mall_user_id = '10x10')  "
			ElseIf FRectSiteName = "itsShinsegaeDept" Then
				wheredetail = wheredetail + " and m.sitename in ('itsShinsegaeDept', 'itsShinsegaeDeptGJ')  "
			Else
				wheredetail = wheredetail + " and m.sitename ='" + FRectSiteName + "'"
			End If

		end if

		if (FRectRdSite<>"") then
			wheredetail = wheredetail + " and m.rdsite ='" + FRectRdSite + "'"
		end if

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if

		if (FRectOnlyOutMall<>"") then
			wheredetail = wheredetail + " and m.accountdiv='50'"
		end if

		if (FRectOnlyPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv='30'"
		end if

        if (FRectIsFlower="Y") then
			wheredetail = wheredetail + " and m.cardribbon is Not NULL "
		end if
		if (FRectIsMinus="Y") then
			wheredetail = wheredetail + " and m.jumundiv='9' "
		end if
		if (FRectIsForeign<>"") then
            wheredetail = wheredetail + " and ((IsNULL(m.dlvcountryCode,'KR')<>'KR') and (IsNULL(m.dlvcountryCode,'KR')<>'ZZ')) "
        end if
		if (FRectIsMilitary<>"") then
            wheredetail = wheredetail + " and (IsNULL(m.dlvcountryCode,'KR') = 'ZZ') "
        end if

        ''2015/12/21
        if (FRectPojangOrder<>"") then
            wheredetail = wheredetail + " and m.ordersheetyn = 'P' "
        end if

        if (FRectJumunDiv<>"") then
            wheredetail = wheredetail + " and m.jumundiv='"&FRectJumunDiv&"'"
        end if

        if (FRectreqHp<>"") then
            wheredetail = wheredetail + " and m.reqhp='"&FRectreqHp&"'"
        end if

        if (FRectBuyHp<>"") then
            wheredetail = wheredetail + " and m.buyhp='"&FRectBuyHp&"'"
        end if

        if (FRectBuyEmail<>"") then
            wheredetail = wheredetail + " and m.buyemail='"&FRectBuyEmail&"'"
        end if

        if (FRectreqdate<>"") then
            wheredetail = wheredetail + " and convert(varchar(10),m.reqdate,121)='"&FRectreqdate&"'"
        end if

        if (FRectSellChannelDiv<>"") then
            wheredetail = wheredetail + " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
        end if

		if (FRectAuthCode<>"") then
			wheredetail = wheredetail + " and m.authcode ='"& FRectAuthCode &"'"
		end if

        ''FRectExcept3pl
        if (FRectExcept3pl<>"") then
			wheredetail = wheredetail + " and m.beadaldiv<>90"
		end if

		if (FRectIsSendGift="Y") then
			wheredetail = wheredetail + " and Exists(select f.orderserial from db_order.dbo.tbl_order_gift_data as f where f.orderserial=m.orderserial) "
		end if

		sqlStr = "select count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal"
		sqlStr = sqlStr + " from " + CStr(tblMasterName) + " m WITH(READUNCOMMITTED) "
		sqlStr = sqlStr & " left join db_order.dbo.tbl_order_custom_number c" & vbcrlf
		sqlStr = sqlStr & " 	on m.orderserial = c.orderserial" & vbcrlf
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + wheredetail
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    		FTotalCount = rsget("cnt")

    		FSubtotal = rsget("subtotal")
    		FAvgTotal = rsget("avgtotal")

    		if IsNull(FSubtotal) then FSubtotal=0
    		if IsNull(FAvgTotal) then FAvgTotal=0
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, convert(varchar,m.regdate,20) as cvreg, convert(varchar,m.beadaldate,20) as cvbeadal"
		if (FRectQryDLVsum<>"") then
		   sqlStr = sqlStr + " ,isNULL((select sum(d.itemcost*d.itemno) from "&tblDetailName&" d with (nolock) where m.orderserial=d.orderserial and d.cancelyn<>'Y' and d.itemid=0),0) as dlvsum"
		end if
		sqlStr = sqlStr & " , c.customnumber" & vbcrlf
		sqlStr = sqlStr + " from " + CStr(tblMasterName) + " m with (nolock) "
		sqlStr = sqlStr & " left join db_order.dbo.tbl_order_custom_number c" & vbcrlf
		sqlStr = sqlStr & " 	on m.orderserial = c.orderserial" & vbcrlf
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by m.idx desc"
		'response.write sqlStr &"<BR>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CLNG(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FMasterItemList(i) = new CJumunMasterItem

				FMasterItemList(i).fcustomnumber = rsget("customnumber")
				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
				FMasterItemList(i).Fuserid		= rsget("userid")
				FMasterItemList(i).Faccountname	= db2Html(rsget("accountname"))
				FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
				FMasterItemList(i).Faccountno	= rsget("accountno")
				FMasterItemList(i).Ftotalvat	= Null2Zoro(rsget("totalvat"))
				FMasterItemList(i).Ftotalmileage= Null2Zoro(rsget("totalmileage"))
				FMasterItemList(i).Ftotalsum	= rsget("totalsum")
				FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
				FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
				FMasterItemList(i).Fregdate		= rsget("cvreg")
				FMasterItemList(i).FCvbeadaldate= rsget("cvbeadal")
				FMasterItemList(i).Fbeadaldiv	= rsget("beadaldiv")
				FMasterItemList(i).Fbeadaldate	= rsget("beadaldate")
				FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
				FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FMasterItemList(i).Fbuyphone	= rsget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsget("buyemail")
				FMasterItemList(i).freqemail	= rsget("reqemail")
				FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
				FMasterItemList(i).Freqzipcode	= rsget("reqzipcode")
				FMasterItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FMasterItemList(i).Freqphone	= rsget("reqphone")
				FMasterItemList(i).Freqhp		= rsget("reqhp")
				FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
				FMasterItemList(i).Fdeliverno	= rsget("deliverno")
				FMasterItemList(i).Fsitename	= rsget("sitename")
				FMasterItemList(i).Fpaygatetid	= rsget("paygatetid")
				FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
				FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))
				FMasterItemList(i).Fresultmsg		= rsget("resultmsg")
				FMasterItemList(i).Frduserid		= rsget("rduserid")
				FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsget("miletotalprice"))
				FMasterItemList(i).Fjungsanflag		= rsget("jungsanflag")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FMasterItemList(i).Fauthcode		= rsget("authcode")
				FMasterItemList(i).Fcouponpay	    = rsget("tencardspend")
				FMasterItemList(i).Fcardribbon	= rsget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsget("fromname"))
				FMasterItemList(i).Freqdate  	= rsget("reqdate")
				FMasterItemList(i).Freqtime 	= db2html(rsget("reqtime"))
				FMasterItemList(i).Fuserlevel	= rsget("userlevel")
				FMasterItemList(i).FsumpaymentEtc  = rsget("sumpaymentEtc")
                FMasterItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

                IF IsNULL(FMasterItemList(i).FsumpaymentEtc) then FMasterItemList(i).FsumpaymentEtc=0

                if (FRectQryDLVsum<>"") then
                    FMasterItemList(i).FdlvPaySum = rsget("dlvsum")
                end if

                FMasterItemList(i).FallAtDiscountPrice = Null2Zoro(rsget("allAtDiscountPrice"))
                FMasterItemList(i).FsubtotalpriceCouponNotApplied = Null2Zoro(rsget("subtotalpriceCouponNotApplied"))
                FMasterItemList(i).FRdsite= rsget("Rdsite")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub getordertemplist()
		dim sqlStr, sqlsearch

		if (FRectRegStart<>"") then
			sqlsearch = sqlsearch & " and t.regdate >='" & CStr(trim(FRectRegStart)) & "'"
		end if
		if (FRectRegEnd<>"") then
			sqlsearch = sqlsearch & " and t.regdate <'" & CStr(trim(FRectRegEnd)) & "'"
		end if
		if FRectorderserial<>"" then
			sqlsearch = sqlsearch & " and t.orderserial='" & trim(FRectorderserial) & "'"
		end if
		if FRectuserid<>"" then
			sqlsearch = sqlsearch & " and t.userid='" & trim(FRectuserid) & "'"
		end if
		if FRectpaygateid<>"" then
			sqlsearch = sqlsearch & " and t.p_tid='" & trim(FRectpaygateid) & "'"
		end if
		if FRectauthcode<>"" then
			sqlsearch = sqlsearch & " and t.p_auth_no='" & trim(FRectauthcode) & "'"
		end if
		if FRectprice<>"" then
			sqlsearch = sqlsearch & " and t.price=" & trim(FRectprice) & ""
		end if
		if FRectbuyname<>"" then
			sqlsearch = sqlsearch & " and t.buyname='" & trim(FRectbuyname) & "'"
		end if
		if FRectbuyhp<>"" then
			sqlsearch = sqlsearch & " and replace(t.buyhp,'-','')='" & replace(trim(FRectbuyhp),"-","") & "'"
		end if
		if FRectreqname<>"" then
			sqlsearch = sqlsearch & " and t.reqname='" & trim(FRectreqname) & "'"
		end if
		if FRectreqhp<>"" then
			sqlsearch = sqlsearch & " and t.reqhp='" & trim(FRectreqhp) & "'"
		end if
		if FRecttemp_idx<>"" then
			sqlsearch = sqlsearch & " and t.temp_idx=" & trim(FRecttemp_idx) & ""
		end if
		if FRectResultMessage<>"" then
			sqlsearch = sqlsearch & " and t.p_rmesg1 like '%" & trim(FRectResultMessage) & "%'"
		end if
		if FRectIsPay<>"" then
			sqlsearch = sqlsearch & " and t.ispay='" & trim(FRectIsPay) & "'"
		end if
		if FRectIsSuccess<>"" then
			sqlsearch = sqlsearch & " and t.issuccess='" & trim(FRectIsSuccess) & "'"
		end if

		sqlStr = "select count(t.temp_idx) as cnt, CEILING(CAST(COUNT(t.temp_idx) AS FLOAT)/" & CStr(FPageSize) & ") as TotalPage "
		if FRectOldJumun="on" then
			sqlStr = sqlStr & " from db_log.dbo.tbl_order_temp_Back t with (nolock)"
		else
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_temp t with (readuncommitted)"
		end if
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    		FTotalCount = rsget("cnt")
			FtotalPage = rsget("TotalPage")
		rsget.Close

		if FTotalCount<1 then exit Sub

		'sqlStr = "select top " & CStr(FPageSize*FCurrPage) & ""
		sqlStr = "select "
		sqlStr = sqlStr & " t.temp_idx, t.regdate,t.userid,t.pggubun,t.tn_paymethod,t.mid,t.p_tid,t.p_auth_no,t.p_rmesg1,t.ispay,t.payresultcode,t.orderserial" & vbcrlf
		sqlStr = sqlStr & " ,t.issuccess,t.price,t.buyname,t.buyhp,t.buyemail,t.reqname,t.reqhp,t.reqemail, t.rdsite" & vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr & " from db_log.dbo.tbl_order_temp_Back t with (nolock)"
		else
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_temp t with (readuncommitted)"
		end if
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by t.temp_idx desc"
		sqlStr = sqlStr & " OFFSET " & CStr((FCurrPage-1)*FPageSize) & " ROWS FETCH NEXT " & CStr(FPageSize) & " ROWS ONLY "

		'response.write sqlStr & "<br>"
		'rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		'FtotalPage =  CLNG(FTotalCount\FPageSize)
		'if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
		'	FtotalPage = FtotalPage +1
		'end if
		'FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			'rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CJumunMasterItem
				FItemList(i).ftemp_idx = rsget("temp_idx")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fpggubun = rsget("pggubun")
				FItemList(i).Faccountdiv = rsget("tn_paymethod")
				FItemList(i).fmid = rsget("mid")
				FItemList(i).ftid = rsget("p_tid")
				FItemList(i).fauth_no = rsget("p_auth_no")
				FItemList(i).fp_rmesg1 = db2html(rsget("p_rmesg1"))
				FItemList(i).fispay = rsget("ispay")
				FItemList(i).fpayresultcode = rsget("payresultcode")
				FItemList(i).forderserial = rsget("orderserial")
				FItemList(i).fissuccess = rsget("issuccess")
				FItemList(i).fprice = rsget("price")
				FItemList(i).fbuyname = db2html(rsget("buyname"))
				FItemList(i).fbuyhp = db2html(rsget("buyhp"))
				FItemList(i).fbuyemail = db2html(rsget("buyemail"))
				FItemList(i).freqname = db2html(rsget("reqname"))
				FItemList(i).freqhp = db2html(rsget("reqhp"))
				FItemList(i).freqemail = db2html(rsget("reqemail"))
				FItemList(i).frdsite = rsget("rdsite")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	' /cscenter/inspect/ordertemplist_giftcard.asp
	public Sub getgiftcardordertemplist()
		dim sqlStr, sqlsearch

		if (FRectRegStart<>"") then
			sqlsearch = sqlsearch & " and t.regdate >='" & CStr(trim(FRectRegStart)) & "'"
		end if
		if (FRectRegEnd<>"") then
			sqlsearch = sqlsearch & " and t.regdate <'" & CStr(trim(FRectRegEnd)) & "'"
		end if
		if FRectgiftOrderSerial<>"" then
			sqlsearch = sqlsearch & " and t.giftOrderSerial='" & trim(FRectgiftOrderSerial) & "'"
		end if
		if FRectuserid<>"" then
			sqlsearch = sqlsearch & " and t.userid='" & trim(FRectuserid) & "'"
		end if
		if FRectpaygateid<>"" then
			sqlsearch = sqlsearch & " and t.p_tid='" & trim(FRectpaygateid) & "'"
		end if
		if FRectauthcode<>"" then
			sqlsearch = sqlsearch & " and t.p_auth_no='" & trim(FRectauthcode) & "'"
		end if
		if FRectprice<>"" then
			sqlsearch = sqlsearch & " and t.price=" & trim(FRectprice) & ""
		end if
		if FRectbuyname<>"" then
			sqlsearch = sqlsearch & " and t.buyname='" & trim(FRectbuyname) & "'"
		end if
		if FRectbuyhp<>"" then
			sqlsearch = sqlsearch & " and replace(t.buyhp,'-','')='" & replace(trim(FRectbuyhp),"-","") & "'"
		end if
		if FRectreqhp<>"" then
			sqlsearch = sqlsearch & " and t.reqhp='" & trim(FRectreqhp) & "'"
		end if
		if FRecttemp_idx<>"" then
			sqlsearch = sqlsearch & " and t.temp_idx=" & trim(FRecttemp_idx) & ""
		end if

		sqlStr = "select count(t.temp_idx) as cnt"
		sqlStr = sqlStr & " from db_order.dbo.tbl_giftcard_order_temp t with (readuncommitted)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount<1 then exit Sub

		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & ""
		sqlStr = sqlStr & " t.temp_idx, t.regdate,t.userid,'' as pggubun,t.mid,t.p_tid,t.p_auth_no,t.p_rmesg1,t.ispay,t.payresultcode,t.giftOrderSerial" & vbcrlf
		sqlStr = sqlStr & " ,t.issuccess,t.price,t.buyname,t.buyhp,t.buyemail,t.reqhp,t.reqemail, t.rdsite" & vbcrlf
		sqlStr = sqlStr & " from db_order.dbo.tbl_giftcard_order_temp t with (readuncommitted)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by t.temp_idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CLNG(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CJumunMasterItem
				FItemList(i).ftemp_idx = rsget("temp_idx")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fpggubun = rsget("pggubun")
				FItemList(i).fmid = rsget("mid")
				FItemList(i).ftid = rsget("p_tid")
				FItemList(i).fauth_no = rsget("p_auth_no")
				FItemList(i).fp_rmesg1 = db2html(rsget("p_rmesg1"))
				FItemList(i).fispay = rsget("ispay")
				FItemList(i).fpayresultcode = rsget("payresultcode")
				FItemList(i).fgiftOrderSerial = rsget("giftOrderSerial")
				FItemList(i).fissuccess = rsget("issuccess")
				FItemList(i).fprice = rsget("price")
				FItemList(i).fbuyname = db2html(rsget("buyname"))
				FItemList(i).fbuyhp = db2html(rsget("buyhp"))
				FItemList(i).fbuyemail = db2html(rsget("buyemail"))
				FItemList(i).freqhp = db2html(rsget("reqhp"))
				FItemList(i).freqemail = db2html(rsget("reqemail"))
				FItemList(i).frdsite = rsget("rdsite")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchMatchJumunList()
		dim sqlStr, wheredetail, i

		wheredetail = ""

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectUserID<>"") then
			wheredetail = wheredetail + " and userid='" + FRectUserID + "'"
		end if

		if (FRectckdate<>"") then
			wheredetail = wheredetail + " and regdate >='" + CStr(FRectRegStart) + "'"
			wheredetail = wheredetail + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectSearchtype01<>"") then
			wheredetail = wheredetail + " and ( accountname like '" + FRectIpkumName + "%'"
			wheredetail = wheredetail + " or buyname = '" + FRectIpkumName + "'"
			wheredetail = wheredetail + " or reqname = '" + FRectIpkumName + "')"
		end if

		if (FRectSearchtype02<>"") and (FRectSubTotalPrice <> "") then
			wheredetail = wheredetail + " and subtotalprice-IsNULL(sumpaymentEtc,0) =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and cancelyn ='N'"
		end if

		if (FRectIpkumDiv2<>"") then
			wheredetail = wheredetail + " and ipkumdiv=2"
		else
			wheredetail = wheredetail + " and ipkumdiv>=2"
		end if

		if (FRectBankOnly <> "") then
			wheredetail = wheredetail + " and accountdiv='7'"
		end if

		sqlStr = "select count(orderserial) as cnt "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master "
		sqlStr = sqlStr + " where idx<>0"
		sqlStr = sqlStr + wheredetail
		''response.write sqlStr
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " *, convert(varchar,regdate,20) as cvreg"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by orderserial desc"

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FMasterItemList(i) = new CJumunMasterItem
				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
				FMasterItemList(i).Fuserid		= rsget("userid")
				FMasterItemList(i).Faccountname	= db2Html(rsget("accountname"))
				FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
				FMasterItemList(i).Faccountno	= rsget("accountno")
				FMasterItemList(i).Ftotalvat	= Null2Zoro(rsget("totalvat"))
				FMasterItemList(i).Ftotalmileage= Null2Zoro(rsget("totalmileage"))
				FMasterItemList(i).Ftotalsum	= rsget("totalsum")
				FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
				FMasterItemList(i).Fipkumdate	= rsget("ipkumdate")
				FMasterItemList(i).Fregdate		= rsget("cvreg")
				FMasterItemList(i).Fbeadaldiv	= rsget("beadaldiv")
				FMasterItemList(i).Fbeadaldate	= rsget("beadaldate")
				FMasterItemList(i).Fcancelyn	= rsget("cancelyn")
				FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FMasterItemList(i).Fbuyphone	= rsget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
				FMasterItemList(i).Freqzipcode	= rsget("reqzipcode")
				FMasterItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FMasterItemList(i).Freqphone	= rsget("reqphone")
				FMasterItemList(i).Freqhp		= rsget("reqhp")
				FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
				FMasterItemList(i).Fdeliverno	= rsget("deliverno")
				FMasterItemList(i).Fsitename	= rsget("sitename")
				FMasterItemList(i).Fpaygatetid	= rsget("paygatetid")
				FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
				FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsget("subtotalprice"))
				FMasterItemList(i).Fresultmsg		= rsget("resultmsg")
				FMasterItemList(i).Frduserid		= rsget("rduserid")
				FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsget("miletotalprice"))
				FMasterItemList(i).Fjungsanflag		= rsget("jungsanflag")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FMasterItemList(i).Fauthcode		= rsget("authcode")
				FMasterItemList(i).Fcouponpay	    = rsget("tencardspend")

				FMasterItemList(i).Fcardribbon	= rsget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsget("fromname"))
				FMasterItemList(i).Freqdate  	= rsget("reqdate")
				FMasterItemList(i).Freqtime 	= db2html(rsget("reqtime"))

				FMasterItemList(i).Fuserlevel	= rsget("userlevel")
                FMasterItemList(i).FsumpaymentEtc  = rsget("sumpaymentEtc")
                IF IsNULL(FMasterItemList(i).FsumpaymentEtc) then FMasterItemList(i).FsumpaymentEtc=0
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'빈 주문데이타(orderserial 이 없는, 빈 디테일 페이지 표시에 사용)
	public Sub GetBlankOneJumunList()
		dim sqlStr, wheredetail, i

		FTotalCount = 1
		FSubtotal=0
		FAvgTotal=0

		FtotalPage =  1
		FResultCount = 1

		redim preserve FMasterItemList(FResultCount)

		set FMasterItemList(i) = new CJumunMasterItem
		FMasterItemList(i).Faccountno	= 0
		FMasterItemList(i).Ftotalvat	= 0
		FMasterItemList(i).Ftotalmileage= 0
		FMasterItemList(i).Ftotalsum	= 0
		FMasterItemList(i).Fdiscountrate	= 1
		FMasterItemList(i).Fsubtotalprice	= 0
		FMasterItemList(i).Fmiletotalprice	= 0

		FMasterItemList(i).Fcouponpay	= 0
	end sub

	public Sub SearchJumunListByDesignerSelllist()
		dim sqlStr, i

		sqlStr = "select "
		sqlStr = sqlStr + " d.itemid, d.itemoption, d.buycash ,d.itemcost, sum(d.itemno) as sm, d.itemname, d.itemoptionname"
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
			sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		end if
		'sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		'sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where d.makerid = '" + FRectDesignerID + "'"
		if (FRectRegStart<>"") then
			if FRectDateType="ipkumil" then
				sqlStr = sqlStr + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			else
				sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if FRectDateType="ipkumil" then
				sqlStr = sqlStr + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			else
				sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectItemid<>"") then
			sqlStr = sqlStr + " and d.itemid=" + FRectItemid
		end if

		if (FRectItemName<>"") then
			sqlStr = sqlStr + " and d.itemname like '" + CStr(FRectItemName) + "%' "
		end if

        sqlStr = sqlStr + " and m.ipkumdiv >3"
		sqlStr = sqlStr + " and m.cancelyn = 'N'"
		sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
		sqlStr = sqlStr + " and d.itemid <> 0"

		sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.buycash, d.itemcost, d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " order by d.itemid desc"

		rsget.PageSize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).Fitemoption   = rsget("itemoption")
				FMasterItemList(i).FItemCost     = rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= rsget("buycash")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    public Sub SearchJumunListByDesignerSelllist_HPDCase()
		dim sqlStr, i

		sqlStr = "select "
		sqlStr = sqlStr + " d.itemid, d.buycash ,d.itemcost, sum(d.itemno) as sm, d.itemname, d.itemoptionname"
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
			sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		end if
		sqlStr = sqlStr + " where d.makerid = '" + FRectDesignerID + "'"
		if (FRectRegStart<>"") then
			if FRectDateType="ipkumil" then
				sqlStr = sqlStr + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			elseif FRectDateType="beasong" then
				sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
				sqlStr = sqlStr + " and d.beasongdate >='" + CStr(FRectRegStart) + "'"
				sqlStr = sqlStr + " and d.currstate=7"
			else
				sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if FRectDateType="ipkumil" then
				sqlStr = sqlStr + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif FRectDateType="beasong" then
				sqlStr = sqlStr + " and d.beasongdate <'" + CStr(FRectRegEnd) + "'"
				sqlStr = sqlStr + " and d.currstate=7"
			else
				sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectItemid<>"") then
			sqlStr = sqlStr + " and d.itemid=" + FRectItemid
		end if

        sqlStr = sqlStr + " and m.ipkumdiv >3"
		sqlStr = sqlStr + " and m.cancelyn = 'N'"
		sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
		sqlStr = sqlStr + " and d.itemid <> 0"

		sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.buycash, d.itemcost, d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " order by d.itemid desc"

		rsget.PageSize = FPageSize
''rw FRectRegEnd
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost     = rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= rsget("buycash")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchJumunListByupcheSelllist2()
		dim sqlStr, wheredetail, i

		wheredetail = ""

		if (FRectRegStart<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			elseif (FRectDateType="beadal") then
				''wheredetail = wheredetail + " and m.beadaldate >='" + CStr(FRectRegStart) + "'"
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
				wheredetail = wheredetail + " and d.beasongdate >='" + CStr(FRectRegStart) + "'" ''상품 배송일로 변경/20120404
			else
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif (FRectDateType="beadal") then
				''wheredetail = wheredetail + " and m.beadaldate <'" + CStr(FRectRegEnd) + "'"
				wheredetail = wheredetail + " and d.beasongdate <'" + CStr(FRectRegEnd) + "'"  ''상품 배송일로 변경/20120404
			else
				wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and m.cancelyn ='N'"
		end if

		if (FRectIpkumDiv4<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv>=4"
		end if

		if (FRectMinusOrderInclude="on") then

		else
		    wheredetail = wheredetail + " and m.jumundiv<>'9'"
		end if

		if (FRectItemid<>"") then
		    wheredetail = wheredetail + " and d.itemid=" + FRectItemid
		end if

		if (FRectDesignerID<>"") then
		    wheredetail = wheredetail + " and d.makerid='" + FRectDesignerID + "'"
		end if

		if FRectCDL<>"" then
			wheredetail = wheredetail + " and i.cate_large='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			wheredetail = wheredetail + " and i.cate_mid='" + FRectCDM + "'"
		end if
		if FRectCDS<>"" then
			wheredetail = wheredetail + " and i.cate_small='" + FRectCDS + "'"
		end if

		If FIsMDPick <> "" Then
			wheredetail = wheredetail + " and d.itemid IN(select linkitemid from [db_sitemaster].[dbo].tbl_main_mdchoice_flash where isNow = 'y') "
		End If

		If FIsMDPickMo <> "" Then
			wheredetail = wheredetail + " and d.itemid IN(select s.itemid from db_sitemaster.dbo.tbl_mobile_main_mdpick_list as l inner join [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item as s on s.listidx = l.idx where l.isusing='Y' and getdate() between l.startdate and l.enddate group by s.itemid) "
		End If

		if (FIsMDPickMoArr<>"") Then
		    wheredetail = wheredetail + " and d.itemid IN(select s.itemid from db_sitemaster.dbo.tbl_mobile_main_mdpick_list as l inner join [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item as s on s.listidx = l.idx where l.isusing='Y' and ((l.startdate >= '"& FRectRegStart &"' and l.startdate < '"& FRectRegEnd &"') or (l.enddate >= '"& FRectRegStart &"' and l.startdate < '"& FRectRegEnd &"')) and s.gubun in ("&FIsMDPickMoArr&") group by s.itemid) "
		end if

		If FIsRdSite <> "" Then
			wheredetail = wheredetail + " and left(m.rdsite,6)='mobile' "
		End If

        if (FRectChannelDiv<>"") then
            wheredetail = wheredetail + " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectChannelDiv)&")"
        end if

'		if (FRectChannelDiv<>"") then
'			if FRectChannelDiv="web" then
'				wheredetail = wheredetail & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
'			elseif FRectChannelDiv="jaehu" then
'				wheredetail = wheredetail & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			elseif FRectChannelDiv="mjaehu" then
'				wheredetail = wheredetail & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			elseif FRectChannelDiv="mobile" then
'				wheredetail = wheredetail & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
'			elseif FRectChannelDiv="ipjum" then
'				wheredetail = wheredetail & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			end if
'		end if

		sqlStr = "select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " d.itemid, d.itemoption, sum(d.itemno) as sm, "
		sqlStr = sqlStr + " d.itemcost,d.buycash,d.itemname,d.itemoptionname"

		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m with (nolock)"
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d with (nolock)"
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
			sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d with (nolock)"
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		end if
		IF (FRectBrandPurchaseType<>"") then
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
		end IF
		if (FRectCDL<>"" or FRectCDM<>"" or FRectCDS<>"") then
		    sqlStr = sqlStr + " Join db_item.dbo.tbl_item as i with (nolock)"
		    sqlStr = sqlStr + " on d.itemid=i.itemid"
		end if
		IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
				sqlStr = sqlStr + "  JOIN db_item.dbo.tbl_display_cate_item as dc with (nolock)"
				sqlStr = sqlStr + " on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF

		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0 and m.ipkumdiv>1"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.itemcost,d.buycash,d.itemname,d.itemoptionname"
		sqlStr = sqlStr + " order by sm desc, d.itemid, d.itemoption"

		rsget.PageSize = FPageSize

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost       = rsget("itemcost")
				FMasterItemList(i).Fbuycash       = rsget("buycash")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchJumunListByDesignerOnlySend()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectUserID<>"") then
			wheredetail = wheredetail + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			wheredetail = wheredetail + " and m.buyname = '" + FRectBuyname + "'"
		end if

		if (FRectReqName<>"") then
			wheredetail = wheredetail + " and m.reqname = '" + FRectReqName + "'"
		end if

		if (FRectSubTotalPrice<>"") then
			wheredetail = wheredetail + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectOnlyIpkumDiv<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv=" + CStr(FRectOnlyIpkumDiv)
		end if

		if (FRectIpkumName<>"") then
			wheredetail = wheredetail + " and m.accountname = '" + FRectIpkumName + "'"
		end if

		if (FRectSiteName<>"") then
			wheredetail = wheredetail + " and m.sitename ='" + FRectSiteName + "'"
		end if

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and m.accountdiv<>'30'"
		end if

		sqlStr = "select"
		sqlStr = sqlStr + " d.orderserial, m.buyname,m.reqname, m.jumundiv, m.userid,"
		sqlStr = sqlStr + " m.ipkumdiv, m.accountdiv, m.regdate, m.reqphone, m.reqhp, m.deliverno, "
		sqlStr = sqlStr + " m.sitename, m.discountrate, m.cancelyn, "
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoption, d.itemoptionname, d.isupchebeasong"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by m.regdate desc"

'response.write sqlStr
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

			set FMasterItemList(i) = new CDesignerJumunList

			FMasterItemList(i).Forderserial = rsget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
			FMasterItemList(i).Fuserid		= rsget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsget("regdate")
			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
			FMasterItemList(i).Freqphone	= rsget("reqphone")
			FMasterItemList(i).Freqhp		= rsget("reqhp")
			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
			FMasterItemList(i).Fdeliverytype	= rsget("isupchebeasong")
			FMasterItemList(i).Fsitename	= rsget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
			FMasterItemList(i).FCancelyn	= rsget("cancelyn")

			FMasterItemList(i).FItemID       = rsget("itemid")
			FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
			FMasterItemList(i).FItemOption   = rsget("itemoption")
			FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))

			rsget.movenext
			i = i + 1
		loop
	end if
		rsget.Close
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
end Class
%>
