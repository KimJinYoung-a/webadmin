<%

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




Class COrderDetailItem
    public Fidx
	public Forderserial
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

    public Fentryname
    public Fentryhp

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

    public FmatcostAdded		'재료비
    public Fmatinclude_yn		'재료비 포함여부

    public FcouponNotAsigncost	'쿠폰/할인 적용이전 원 강좌료

    public Frefundstate			'환불상태

    ''쿠폰 적용 주문인지 체크
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
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

	Public function GetRefundStateName()
        if Frefundstate="8" then
            GetRefundStateName = "재료비 80%"
	    elseif Frefundstate="9" then
		    GetRefundStateName = "재료비 90%"
	    elseif Frefundstate="0" then
		    GetRefundStateName = "전액환불"
	    else
		    if (FCancelYn = "Y") then
		    	GetRefundStateName = "취소환불"
		    else
		    	GetRefundStateName = "환불없음"
		    end if
	    end if
	 end Function

	public function GetRefundStateColor()
        if Frefundstate="8" then
            GetRefundStateColor = "#CC9933"
	    elseif Frefundstate="9" then
		    GetRefundStateColor = "#CC9933"
	    elseif Frefundstate="0" then
		    GetRefundStateColor = "#FF0000"
	    else
		    if (FCancelYn = "Y") then
		    	GetRefundStateColor = "red"
		    else
		    	GetRefundStateColor = "#000000"
		    end if
	    end if
	end function

	public function GetRefundPrice()
        if (Fmatinclude_yn = "C") then
        	'재료비 포함인경우
	        if Frefundstate="8" then
	            GetRefundPrice = FormatNumber(FmatcostAdded * 0.8, 0)
		    elseif Frefundstate="9" then
			    GetRefundPrice = FormatNumber(FmatcostAdded * 0.9, 0)
		    elseif Frefundstate="0" then
			    GetRefundPrice = FreducedPrice
		    else
			    if (FCancelYn = "Y") then
			    	GetRefundPrice = FreducedPrice
			    else
			    	GetRefundPrice = 0
			    end if
			end if
        else
        	'재료비 별도인경우
        	if Frefundstate="0" then
        		GetRefundPrice = FreducedPrice - FmatcostAdded
        	else
			    if (FCancelYn = "Y") then
			    	GetRefundPrice = FreducedPrice - FmatcostAdded
			    else
			    	GetRefundPrice = 0
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

	public Fgoodsname
	public Fusercnt

	public Fallatdiscountprice

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
    
    public FweClassYn           '''2012 추가
	public FsumPaymentEtc
	public FPgGubun
	
	public function isWeClass() ''단체 강좌인지 여부
	    if isNULL(FweClassYn) then
	        isWeClass = FALSE
	        Exit function
	    end if
	    
	    isWeClass = (FweClassYn="Y")
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
        IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR") and (FDlvcountryCode<>"ZZ")
    end function

    ''군부대배송
    public function IsArmiDeliver()
        IsArmiDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode="ZZ")
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

	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#f0ca2c"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#a3cf6c"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#6ca54e"   ''VIP
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#f68d3f"   ''오렌지
		elseif Fuserlevel="5" then
			GetUserLevelColor = "#865e25"  '' 옐로우
		elseif Fuserlevel="6" then
			GetUserLevelColor = "#B70606"  '' staff
		else
			GetUserLevelColor = "#f0ca2c"
		end if
	end function

	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Seed"
		elseif Fuserlevel="2" then
			GetUserLevelName = "Bud"
		elseif Fuserlevel="3" then
			GetUserLevelName = "Leaf"
		elseif Fuserlevel="4" then
			GetUserLevelName = "Bean"
	    elseif Fuserlevel="5" then
			GetUserLevelName = "Tree"
		elseif Fuserlevel="6" then
			GetUserLevelName = "STAFF"
		else
			GetUserLevelName = "Seed"
		end if
	end function

	public function GetJumunDivName()
		if Fjumundiv="1" then
			GetJumunDivName = "웹주문"
		elseif Fjumundiv="3" then
			GetJumunDivName = "예약주문"
		elseif Fjumundiv="5" then
			GetJumunDivName = "외부몰"
		elseif Fjumundiv="6" then
			GetJumunDivName = "아카데미DIY상품"
		elseif Fjumundiv="7" then
			GetJumunDivName = "플라워"
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
			IpkumDivColor="#FF00FF"
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
		elseif Faccountdiv="900" then
			JumunMethodName="수기"
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
			IpkumDivName="결제대기"
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

    public FRectIsWeClass
    
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
		''갯수
		sqlStr = "select count(*) as cnt "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
    		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
    		IF (FRectIsWeClass="Y") then
    		    sqlStr = sqlStr + " Join [db_academy].[dbo].tbl_academy_order_detail dd"
    		    sqlStr = sqlStr + " on m.orderserial=dd.orderserial"
    		    sqlStr = sqlStr + " and dd.weClassyn='Y'"
    		END IF
    	end if
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and sitename <> 'diyitem'"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

        if (FRectIsForeign<>"") then
            sqlStr = sqlStr + " and IsNULL(dlvcountryCode,'KR')<>'KR'"
        end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname = '" + FRectReqName + "'"  ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname = '" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and deliverno='" + FRectReqSongjangNo + "'"
		end if

		if (FRectIsFlower="Y") then
			sqlStr = sqlStr + " and cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			sqlStr = sqlStr + " and ((reqzipaddr='') or (reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			sqlStr = sqlStr + " and jumundiv='9' "
		end if

		if (FRectExtSiteName<>"") then
			sqlStr = sqlStr + " and ((sitename='" + FRectExtSiteName + "') or (rdsite='" + FRectExtSiteName + "')) "
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "

		sqlStr = sqlStr & " ,IsNULL((select count(itemid) as cnt from [db_academy].[dbo].tbl_academy_order_detail D where D.orderserial=m.orderserial and  D.cancelyn <> 'Y' ),0) as cnt"
		sqlStr = sqlStr & " ,isNULL((select top 1 D1.weClassyn from [db_academy].[dbo].tbl_academy_order_detail D1 where D1.orderserial=m.orderserial),'N') as weClassyn"  
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		IF (FRectIsWeClass="Y") then
		    sqlStr = sqlStr + " Join [db_academy].[dbo].tbl_academy_order_detail dd"
		    sqlStr = sqlStr + " on m.orderserial=dd.orderserial"
		    sqlStr = sqlStr + " and dd.weClassyn='Y'"
		END IF
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and sitename <> 'diyitem'"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

        if (FRectIsForeign<>"") then
            sqlStr = sqlStr + " and IsNULL(m.dlvcountryCode,'KR')<>'KR'"
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
			sqlStr = sqlStr + " and m.accountname = '" + FRectIpkumName + "'" ''like
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

		if (FRectIsFlower="Y") then
			sqlStr = sqlStr + " and m.cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			sqlStr = sqlStr + " and ((m.reqzipaddr='') or (reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			sqlStr = sqlStr + " and m.jumundiv='9' "
		end if

		if (FRectExtSiteName<>"") then
			sqlStr = sqlStr + " and ((m.sitename='" + FRectExtSiteName + "') or (m.rdsite='" + FRectExtSiteName + "')) "
		end if

        'if (FRectBuyname<>"") or (FRectReqName<>"") or (FRectIpkumName<>"") or (FRectSubTotalPrice<>"") or (FRectBuyHp<>"") or (FRectReqHp<>"") or (FRectBuyPhone<>"") or (FRectReqPhone<>"") or (FRectReqSongjangNo<>"") then
        'sqlStr = sqlStr + " order by orderserial desc"
        'else
		sqlStr = sqlStr + " order by m.idx desc"
	    'end if
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderMasterItem
				FItemList(i).Forderserial       = rsget("orderserial")
				FItemList(i).Fjumundiv	        = rsget("jumundiv")
				FItemList(i).Fuserid			= rsget("userid")
				FItemList(i).Faccountname		= db2Html(rsget("accountname"))
				FItemList(i).Faccountdiv		= trim(rsget("accountdiv"))
				FItemList(i).Faccountno	        = rsget("accountno")

				FItemList(i).Ftotalmileage      = rsget("totalmileage")
				FItemList(i).Ftotalsum	        = rsget("totalsum")
				FItemList(i).Fipkumdiv	        = rsget("ipkumdiv")
				FItemList(i).Fipkumdate	        = rsget("ipkumdate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fbeadaldate		= rsget("beadaldate")
				FItemList(i).Fcancelyn	        = rsget("cancelyn")

				FItemList(i).Fbuyname			= db2Html(rsget("buyname"))
				FItemList(i).Fbuyphone	        = rsget("buyphone")
				FItemList(i).Fbuyhp				= rsget("buyhp")
				FItemList(i).Fbuyemail	        = rsget("buyemail")
				FItemList(i).Freqname			= db2Html(rsget("reqname"))

				FItemList(i).Freqzipcode		= rsget("reqzipcode")
				FItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress		= db2Html(rsget("reqaddress"))
				FItemList(i).Freqphone	        = rsget("reqphone")
				FItemList(i).Freqhp				= rsget("reqhp")
				FItemList(i).Freqemail	        = rsget("reqemail")
				FItemList(i).Fcomment			= db2Html(rsget("comment"))

				FItemList(i).Fdeliverno	        = rsget("deliverno")

				FItemList(i).Fsitename	        = rsget("sitename")
				FItemList(i).Fpaygatetid		= rsget("paygatetid")
				FItemList(i).Fdiscountrate		= rsget("discountrate")
				FItemList(i).Fsubtotalprice		= rsget("subtotalprice")
				FItemList(i).Fresultmsg			= rsget("resultmsg")
				FItemList(i).Frduserid			= rsget("rduserid")
				FItemList(i).Fmiletotalprice	= rsget("miletotalprice")
				if IsNULL(FItemList(i).Fmiletotalprice) then FItemList(i).Fmiletotalprice=0

				FItemList(i).Fauthcode		        = rsget("authcode")
				FItemList(i).Ftencardspend			= rsget("tencardspend")
				FItemList(i).Fuserlevel		        = rsget("userlevel")
				FItemList(i).Fspendmembership		= rsget("spendmembership")

                FItemList(i).Fallatdiscountprice 	= rsget("allatdiscountprice")

                FItemList(i).Freqdate    = rsget("reqdate")
                FItemList(i).Freqtime    = rsget("reqtime")
                FItemList(i).Fcardribbon = rsget("cardribbon")
                FItemList(i).Fmessage    = rsget("message")
                FItemList(i).Ffromname   = rsget("fromname")

                FItemList(i).Fgoodsname  = db2Html(rsget("goodsnames"))
                FItemList(i).Fusercnt    = rsget("cnt")

                FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

                if (IsNull(FItemList(i).Fallatdiscountprice) = true) then
                	FItemList(i).Fallatdiscountprice = 0
                end if
                
                FItemList(i).FweClassyn = rsget("weClassyn")
                
                ''2016/09/09
                FItemList(i).Frdsite	= rsget("rdsite")
                FItemList(i).FsumPaymentEtc = rsget("sumPaymentEtc")
                FItemList(i).FPgGubun       = rsget("pggubun")
    
                if isNULL(FItemList(i).FsumPaymentEtc) then FItemList(i).FsumPaymentEtc=0
                if isNULL(FItemList(i).FPgGubun) then FItemList(i).FPgGubun=""
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub



	public Sub QuickSearchOrderMaster()
		dim sqlStr, i

		sqlStr = "select top 1 m.* "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
		    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		end if
		sqlStr = sqlStr + " where m.idx<>0"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname = '" + FRectReqName + "'" ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname ='" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and deliverno='" + FRectReqSongjangNo + "'"
		end if

		sqlStr = sqlStr + " order by orderserial desc"
        ''sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if not rsget.Eof then
		        FTotalCount = 1
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        if not rsget.Eof then
	        set FOneItem = new COrderMasterItem

			FOneItem.Forderserial           = rsget("orderserial")
			FOneItem.Fjumundiv	            = rsget("jumundiv")
			FOneItem.Fuserid		        = rsget("userid")
			FOneItem.Faccountname	        = db2Html(rsget("accountname"))
			FOneItem.Faccountdiv	        = trim(rsget("accountdiv"))
			FOneItem.Faccountno	            = rsget("accountno")

			FOneItem.Ftotalmileage          = rsget("totalmileage")
			FOneItem.Ftotalsum	            = rsget("totalsum")
			FOneItem.Fipkumdiv	            = rsget("ipkumdiv")
			FOneItem.Fipkumdate	            = rsget("ipkumdate")
			FOneItem.Fregdate		        = rsget("regdate")
			FOneItem.Fbaljudate		        = rsget("baljudate")
			FOneItem.Fbeadaldate	        = rsget("beadaldate")
			FOneItem.Fcancelyn	            = rsget("cancelyn")
			FOneItem.Fbuyname		        = db2Html(rsget("buyname"))
			FOneItem.Fbuyphone	            = rsget("buyphone")
			FOneItem.Fbuyhp		            = rsget("buyhp")
			FOneItem.Fbuyemail	            = rsget("buyemail")
			FOneItem.Freqname		        = db2Html(rsget("reqname"))
			FOneItem.Freqzipcode	        = rsget("reqzipcode")
			FOneItem.Freqaddress	        = db2Html(rsget("reqaddress"))
			FOneItem.Freqphone	            = rsget("reqphone")
			FOneItem.Freqhp		            = rsget("reqhp")
			FOneItem.Freqemail	            = rsget("reqemail")
			FOneItem.Fcomment		        = db2Html(rsget("comment"))
			FOneItem.Fdeliverno	            = rsget("deliverno")
			FOneItem.Fsitename	            = rsget("sitename")
			FOneItem.Fpaygatetid	        = rsget("paygatetid")
			FOneItem.Fdiscountrate	        = rsget("discountrate")
			FOneItem.Fsubtotalprice	        = rsget("subtotalprice")
			FOneItem.Fresultmsg		        = rsget("resultmsg")
			FOneItem.Frduserid		        = rsget("rduserid")
			FOneItem.Fmiletotalprice	    = rsget("miletotalprice")

			FOneItem.FInsureCd           	= rsget("InsureCd")

			if IsNULL(FOneItem.Fmiletotalprice) then FOneItem.Fmiletotalprice=0

			FOneItem.Fjungsanflag		    = rsget("jungsanflag")
			FOneItem.Freqzipaddr		    = db2Html(rsget("reqzipaddr"))
			FOneItem.Fauthcode		        = rsget("authcode")
			FOneItem.Fcashreceiptreq		= rsget("cashreceiptreq")

			FOneItem.Ftencardspend		    = rsget("tencardspend")

			FOneItem.Fuserlevel		        = rsget("userlevel")
			FOneItem.Fspendmembership	    = rsget("spendmembership")
			FOneItem.Fallatdiscountprice    = rsget("allatdiscountprice")

			if IsNULL(FOneItem.Fspendmembership) then FOneItem.Fspendmembership=0
			if IsNULL(FOneItem.Fallatdiscountprice) then FOneItem.Fallatdiscountprice=0

			FOneItem.Freqdate    = rsget("reqdate")
            FOneItem.Freqtime    = rsget("reqtime")
            FOneItem.Fcardribbon = rsget("cardribbon")
            FOneItem.Fmessage    = rsget("message")
            FOneItem.Ffromname   = rsget("fromname")

            FOneItem.FDlvcountryCode = rsget("DlvcountryCode")
            FOneItem.Frdsite	= rsget("rdsite")

	    end if
		rsget.Close

		if (FResultCount>0) then
    		if (FOneItem.Faccountdiv="110") then
    		    sqlStr = "select IsNULL(sum(acctamount),0) as okcashbagSpend"
    			sqlStr = sqlStr + "	from db_order.dbo.tbl_order_paymentEtc"
    			sqlStr = sqlStr + "	where orderserial='"&FRectOrderSerial&"'"
    			sqlStr = sqlStr + "	and acctdiv='110'"
    			rsget.Open sqlStr,dbget,1
    			if not rsget.Eof then
    		        FOneItem.FokcashbagSpend = rsget("okcashbagSpend")
    		    end if
    		    rsget.close
    		end if
    	end if
	end sub

	public Sub QuickSearchOrderDetail()
		dim sqlStr
		dim i

		sqlStr = "select d." & FIELD_DETAILIDX & " as idx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.reducedPrice, d.entryname, d.entryhp, d.matcostAdded, d.matinclude_yn, d.couponNotAsigncost, d.refundstate"
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, i.listimg as listimage "
		sqlStr = sqlStr + " ,i.smallimg as smallimage , i.lec_cost as orgprice, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem, d.requiredetail  "
		sqlStr = sqlStr + " ,d.issailitem, d.bonuscouponidx, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx "
		sqlStr = sqlStr + " ,s.divname as songjangdivname, s.findurl"
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d "
		end if
		sqlStr = sqlStr + "     left join [db_academy].[dbo].tbl_lec_item i on d.itemid=i.idx"
		sqlStr = sqlStr + "     left join " & TABLE_SONGJANG_DIV & " s on d.songjangdiv=s.divcd"
		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"

        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderDetailItem

			FItemList(i).Forderserial = CStr(FRectOrderSerial)
			FItemList(i).Fidx         = rsget("idx")
			FItemList(i).Fmakerid     = rsget("makerid")
			FItemList(i).Fitemid      = rsget("itemid")
			FItemList(i).Fitemoption  = rsget("itemoption")
			FItemList(i).Fitemno      = rsget("itemno")
			FItemList(i).Fitemcost    = rsget("itemcost")
			FItemList(i).Fmileage     = rsget("mileage")
			FItemList(i).Fcancelyn    = rsget("cancelyn")

			FItemList(i).FItemName    = db2html(rsget("itemname"))
			FItemList(i).FSmallImage	= webImgUrl & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")

			if IsNull(rsget("itemoptionname")) then
				FItemList(i).FItemoptionName = "-"
			else
				FItemList(i).FItemoptionName = db2html(rsget("itemoptionname"))
			end if

			FItemList(i).Fcurrstate         = rsget("currstate")
			FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
			FItemList(i).Fsongjangno        = rsget("songjangno")
			FItemList(i).Fbeasongdate       = rsget("beasongdate")
			FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")
			FItemList(i).Fissailitem        = rsget("issailitem")
			FItemList(i).Fupcheconfirmdate    = rsget("upcheconfirmdate")

			FItemList(i).Frequiredetail    = rsget("requiredetail")
            FItemList(i).Fsongjangdivname  = db2html(rsget("songjangdivname"))
            FItemList(i).Ffindurl          = db2html(rsget("findurl"))

            FItemList(i).Fentryname     	= rsget("entryname")
            FItemList(i).Fentryhp     		= rsget("entryhp")

            FItemList(i).FmatcostAdded     		= rsget("matcostAdded")
            FItemList(i).Fmatinclude_yn   		= rsget("matinclude_yn")
            FItemList(i).FcouponNotAsigncost   	= rsget("couponNotAsigncost")

            FItemList(i).Forgprice          = rsget("orgprice")
            FItemList(i).Fissailitem        = rsget("issailitem")
            FItemList(i).Fbonuscouponidx    = rsget("bonuscouponidx")
            FItemList(i).Fitemcouponidx     = rsget("itemcouponidx")

            FItemList(i).FreducedPrice      = rsget("reducedPrice")

            FItemList(i).Frefundstate      	= rsget("refundstate")

            if Not IsNULL(FItemList(i).Fsongjangno) then
               FItemList(i).Fsongjangno = replace(FItemList(i).Fsongjangno,"-","")
            end if
			rsget.movenext
			i=i+1
		loop
		rsget.close
	end sub

    public function GetOneOrderDetail
        dim sqlStr, i
	    dim mastertable, detailtable

	    if (FRectOldOrder<>"") then
			mastertable = "[db_log].[dbo].tbl_old_order_master_2003"
			detailtable	= "[db_log].[dbo].tbl_old_order_detail_2003"
		else
			mastertable = "[db_academy].[dbo].tbl_academy_order_master"
			detailtable	= "" & TABLE_ORDERDETAIL & ""
		end if

		sqlStr =	" SELECT d.idx, d.itemid, d.itemoption, d.itemno, d.itemoptionname, d.itemcost," &_
					" d.itemname, d.itemcost, d.makerid, d.currstate, replace(d.songjangno,'-','') as songjangno, d.songjangdiv," &_
					" d.cancelyn, d.isupchebeasong, d.mileage, d.requiredetail, d.oitemdiv, d.beasongdate, d.issailitem, d.upcheconfirmdate," &_
					" d.bonuscouponidx, d.itemcouponidx, d.reducedPrice," &_
					" i.smallimg as smallimage, i.listimg as listimage, i.brandname, i.itemdiv, i.lec_cost as orgprice" &_
					" ,s.divname,s.findurl ,s.tel as DeliveryTel" &_
					" FROM " + detailtable + " d " &_
					" JOIN [db_academy].[dbo].tbl_lec_item i" &_
					"		ON d.itemid=i.idx " &_
					" LEFT JOIN " & TABLE_SONGJANG_DIV & " s " &_
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
			FItemList(i).FSmallImage	= webImgUrl & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")

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