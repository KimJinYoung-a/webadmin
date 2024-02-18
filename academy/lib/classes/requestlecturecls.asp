<%

'[db_academy].[dbo].tbl_academy_order_master
'[db_academy].[dbo].tbl_academy_order_detail
'
'jumundiv   1   DIY주문
'jumundiv   3   예약주문
'jumundiv   8   강좌신청
'
'여기서는 강좌신청만 다룬다.
'디테일에 들어있는 강좌정보를 마스터테이블에 붙여서 구성한다.

Class CRequestLectureDetailItem
    public Fdetailidx
    public Fmasteridx
    public Forderserial

    public Fentryname
    public Fentryhp
    public Fcancelyn

    public Fitemoptionname

	public FMakerid
	public Foitemdiv
	public Fitemid
	public Fitemoption
	public Fitemno
	public Fitemcost
	public Fbuycash
	'public Fitemvat
	public Fmileage
	''public Fcosttotal
	''public Forderdate

	public FItemName
	public FImageList
	public FImageSmall

    public Fcurrstate
    public Fsongjangdiv
    public Fsongjangno
    public Fupcheconfirmdate
    public Fbeasongdate

    public Fisupchebeasong
    public Fissailitem

    public Frequiredetail

    public Fsongjangdivname
    public Ffindurl

    public Freducedprice
    public FmatcostAdded

    public Fbonuscouponidx
    public Fitemcouponidx

    public function CancelStateStr()
        if (FCancelYn="D") then
            CancelStateStr = "삭제"
        elseif (FCancelYn="Y") then
            CancelStateStr = "취소"
        elseif (FCancelYn="A") then
            CancelStateStr = "추가"
        else
            CancelStateStr = "정상"
        end if
    end function

    public function CancelStateColor()
        if (FCancelYn="D") then
            CancelStateColor = "#FF0000"
        elseif (FCancelYn="Y") then
            CancelStateColor = "#FF0000"
        else
            CancelStateColor = "#000000"
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

    ''주문제작 상품
    public function IsRequireDetailExistsItem()
        IsRequireDetailExistsItem = (Foitemdiv="06") or (Frequiredetail<>"")
    end function

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
		else
			BeasongCD2Name = "일반택배"
		end if
	end function

    Private Sub Class_Initialize()
        '
    End Sub

    Private Sub Class_Terminate()
        '
    End Sub
end Class

Class CRequestLectureMasterItem
    public Fidx
    public Forderserial
    public Fjumundiv
    public Fuserid
    public Faccountname
    public Faccountdiv
    public Ftotalitemno
    public Ftotalmileage
    public Ftotalsum
    public Fdiscountrate
    public Fdiscountprice
    public Fcancelitemno
    public Fcancelprice
    public Fsubtotalitemno
    public Fsubtotalprice
    public Fipkumdiv
    public Fipkumdate
    public Fregdate
    public Fbeadaldate
    public Fbaljudate
    public Fcanceldate
    public Fcancelyn
    public Fbuyname
    public Fbuyphone
    public Fbuyhp
    public Fbuyemail
    public Freqname
    public Freqzipcode
    public Freqzipaddr
    public Freqaddress
    public Freqphone
    public Freqhp
    public Freqemail
    public Fcomment
    public Fsitename
    public Fpaygatetid
    public Fresultmsg
    public Frduserid
    public Fmilelogid
    public Fmiletotalprice
    public Fjungsanflag
    public Fauthcode
    public Frdsite
    public Ftencardspend
    public Fbeasongmemo
    public Freqdate
    public Freqtime
    public Fcardribbon
    public Fmessage
    public Ffromname
    public Fcashreceiptreq
    public Finireceipttid
    public Freferip
    public Fuserlevel
    public Flinkorderserial
    public Fspendmembership
    public Fsentenceidx
    public Freguserid
    public Foldorderserial

	public Fgoodsnames

    '강좌정보
    public Fitemcost
    public Fbuycash
    public Fitemname
    public FImageList
    public FImageSmall
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fmakername
    public Fitemoptionname
    public Fbeasongdate
	public Fmileage
	public Fitemno

	public Fjupsustartday
	public Fjupsuendday
	public Flecturestartday
	public Flectureendday
	public Flimitmaxno
	public Flimitminno
	public Flimitsoldno
	public Flimitwaitno

    ''단체강좌
    public FWeClassYN
    public FWantStudyName
    public FWantStudyYear
    public FWantStudyMonth
    public FWantStudyDay
    public FWantStudyAmPm
    public FWantStudyHour
    public FWantStudyMin
    public FWantStudyPlace
    public FWantStudyWho
	'public F


	public function GetIpkumDivColor()
        GetIpkumDivColor = LectureIpkumDivColor(FIpkumDiv)
	end function

	public function GetIpkumDivName()
		GetIpkumDivName = LectureIpkumDivName(FIpkumDiv)
	end function

	public function GetJumunDivName()
		if Fjumundiv="1" then
			GetJumunDivName = "DIY주문"
		elseif Fjumundiv="3" then
			GetJumunDivName = "예약주문"
		elseif Fjumundiv="8" then
			GetJumunDivName = "강좌신청"
		else
			GetJumunDivName = Fjumundiv
		end if
	end function

    public function IsAvailable()
        IsAvailable = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
    end function

    public function GetUserLevelColor()
        if Fuserlevel="1" then
            GetUserLevelColor = "#44EE44"
        elseif Fuserlevel="2" then
            GetUserLevelColor = "#4444EE"
        elseif Fuserlevel="3" then
            GetUserLevelColor = "#EE4444"
        elseif Fuserlevel="9" then
            GetUserLevelColor = "#FF44FF"  ''magenta
        else
            GetUserLevelColor = "#000000"
        end if
    end function

    public function GetUserLevelName()
        if Fuserlevel="1" then
            GetUserLevelName = "Green"
        elseif Fuserlevel="2" then
            GetUserLevelName = "Blue"
        elseif Fuserlevel="3" then
            GetUserLevelName = "VIP"
        elseif Fuserlevel="9" then
            GetUserLevelName = "Mania"  ''magenta
        else
            GetUserLevelName = "Yellow"
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
        elseif FCancelYn="Y" then
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
            IpkumDivColor="#FFFF00"
        elseif Fipkumdiv="7" then
            IpkumDivColor="#EE2222"
        elseif Fipkumdiv="8" then
            IpkumDivColor="#FF00FF"
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
        else
            JumunMethodName = "-"
        end if
    end function

    Public function IpkumDivName()
    	if (Fsitename = "academy") then

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
	            IpkumDivName="강좌준비"
	        elseif Fipkumdiv="6" then
	            IpkumDivName="강좌확정"
	        elseif Fipkumdiv="7" then
	            IpkumDivName="강좌확정"
	        else
	            IpkumDivName=Fipkumdiv
	        end if
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
				IpkumDivName="상품출고"
			end if
		end if
    end Function

    public function SiteNameColor()
        if Fsitename<>"academy" then
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

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CRequestLecture
    public FOneItem
    public FItemList()

    public FTotalCount
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount

    public FRectidx
    public FRectRegStart
    public FRectRegEnd
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
    public FRectItemID
    public FRectSiteName
    public FRectOldjumun            'deprecated

    public Sub GetRequestLectureMasterOne()
        dim sqlStr,i

        sqlStr = " select top 1 orderserial from [db_academy].[dbo].tbl_academy_order_master m "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.jumundiv in ('8', '1') "
        sqlStr = sqlStr + " and m.orderserial='" + FRectOrderSerial + "' "

        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        if  not rsACADEMYget.EOF  then
            FTotalCount = 1
        else
            FTotalCount = 0
        end if
        rsACADEMYget.Close

        sqlStr = " select top 1 "
        sqlStr = sqlStr + " m.* from [db_academy].[dbo].tbl_academy_order_master m "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.jumundiv in ('8', '1') "
        sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "' "
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        'response.write sqlStr
        FResultCount = FTotalCount

        set FOneItem = new CRequestLectureMasterItem

        if  not rsACADEMYget.EOF  then
            FOneItem.Fidx              = rsACADEMYget("idx")
            FOneItem.Forderserial      = rsACADEMYget("orderserial")
            FOneItem.Fjumundiv         = Trim(rsACADEMYget("jumundiv"))
            FOneItem.Fuserid           = rsACADEMYget("userid")
            FOneItem.Faccountname      = db2html(rsACADEMYget("accountname"))
            FOneItem.Faccountdiv       = Trim(CStr(rsACADEMYget("accountdiv")))
            FOneItem.Ftotalitemno      = rsACADEMYget("totalitemno")
            FOneItem.Ftotalmileage     = rsACADEMYget("totalmileage")
            FOneItem.Ftotalsum         = rsACADEMYget("totalsum")
            FOneItem.Fdiscountrate     = rsACADEMYget("discountrate")
            FOneItem.Fdiscountprice    = rsACADEMYget("discountprice")
            FOneItem.Fcancelitemno     = rsACADEMYget("cancelitemno")
            FOneItem.Fcancelprice      = rsACADEMYget("cancelprice")
            FOneItem.Fsubtotalitemno   = rsACADEMYget("subtotalitemno")
            FOneItem.Fsubtotalprice    = rsACADEMYget("subtotalprice")
            FOneItem.Fipkumdiv         = rsACADEMYget("ipkumdiv")
            FOneItem.Fipkumdate        = rsACADEMYget("ipkumdate")
            FOneItem.Fregdate          = rsACADEMYget("regdate")
            FOneItem.Fbeadaldate       = rsACADEMYget("beadaldate")
            FOneItem.Fbaljudate        = rsACADEMYget("baljudate")
            FOneItem.Fcanceldate       = rsACADEMYget("canceldate")
            FOneItem.Fcancelyn         = UCase(rsACADEMYget("cancelyn"))
            FOneItem.Fbuyname          = db2html(rsACADEMYget("buyname"))
            FOneItem.Fbuyphone         = rsACADEMYget("buyphone")
            FOneItem.Fbuyhp            = rsACADEMYget("buyhp")
            FOneItem.Fbuyemail         = db2html(rsACADEMYget("buyemail"))
            FOneItem.Freqname          = db2html(rsACADEMYget("reqname"))
            FOneItem.Freqzipcode       = rsACADEMYget("reqzipcode")
            FOneItem.Freqzipaddr       = db2html(rsACADEMYget("reqzipaddr"))
            FOneItem.Freqaddress       = db2html(rsACADEMYget("reqaddress"))
            FOneItem.Freqphone         = rsACADEMYget("reqphone")
            FOneItem.Freqhp            = rsACADEMYget("reqhp")
            FOneItem.Freqemail         = db2html(rsACADEMYget("reqemail"))
            FOneItem.Fcomment          = db2html(rsACADEMYget("comment"))
            FOneItem.Fsitename         = rsACADEMYget("sitename")
            FOneItem.Fpaygatetid       = rsACADEMYget("paygatetid")
            FOneItem.Fresultmsg        = rsACADEMYget("resultmsg")
            FOneItem.Frduserid         = rsACADEMYget("rduserid")
            FOneItem.Fmilelogid        = rsACADEMYget("milelogid")
            FOneItem.Fmiletotalprice   = rsACADEMYget("miletotalprice")
            FOneItem.Fjungsanflag      = rsACADEMYget("jungsanflag")
            FOneItem.Fauthcode         = rsACADEMYget("authcode")
            FOneItem.Frdsite           = rsACADEMYget("rdsite")
            FOneItem.Ftencardspend     = rsACADEMYget("tencardspend")
            FOneItem.Fbeasongmemo      = db2html(rsACADEMYget("beasongmemo"))
            FOneItem.Freqdate          = rsACADEMYget("reqdate")
            FOneItem.Freqtime          = rsACADEMYget("reqtime")
            FOneItem.Fcardribbon       = rsACADEMYget("cardribbon")
            FOneItem.Fmessage          = db2html(rsACADEMYget("message"))
            FOneItem.Ffromname         = db2html(rsACADEMYget("fromname"))
            FOneItem.Fcashreceiptreq   = rsACADEMYget("cashreceiptreq")
            FOneItem.Finireceipttid    = rsACADEMYget("inireceipttid")
            FOneItem.Freferip          = rsACADEMYget("referip")
            FOneItem.Fuserlevel        = rsACADEMYget("userlevel")
            FOneItem.Flinkorderserial  = rsACADEMYget("linkorderserial")
            FOneItem.Fspendmembership  = rsACADEMYget("spendmembership")
            FOneItem.Fsentenceidx      = rsACADEMYget("sentenceidx")
            FOneItem.Freguserid        = rsACADEMYget("reguserid")
            FOneItem.Foldorderserial   = rsACADEMYget("oldorderserial")
        end if
        rsACADEMYget.close

        'TODO : 디테일정보에는 하나의 강좌정보만 있다고 가정한다.
        sqlStr = " select top 1 d.*, l.listimg, l.smallimg, l.lecturer_name, o.regstartdate, o.regenddate, o.lecstartdate, o.lecenddate, o.limit_count, o.min_count, o.limit_sold, o.wait_count "

        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_order_detail d "
        sqlStr = sqlStr + " 	join [db_academy].[dbo].tbl_lec_item l "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		d.itemid = l.idx "
        sqlStr = sqlStr + " 	join [db_academy].dbo.tbl_lec_item_option o "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		d.itemid = o.lecidx and d.itemoption = o.lecoption "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and d.orderserial = '" + CStr(FRectOrderSerial) + "' "
        'response.write sqlStr
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        if  not rsACADEMYget.EOF  then

            FOneItem.Fitemcost      = rsACADEMYget("itemcost")
            FOneItem.Fbuycash       = rsACADEMYget("buycash")
            FOneItem.Fitemname      = db2html(rsACADEMYget("itemname"))
            FOneItem.Fitemid        = rsACADEMYget("itemid")
            FOneItem.Fitemoption    = rsACADEMYget("itemoption")
            FOneItem.Fmakerid       = rsACADEMYget("makerid")
            FOneItem.Fmakername     = db2html(rsACADEMYget("lecturer_name"))
            FOneItem.Fitemoptionname= db2html(rsACADEMYget("itemoptionname"))
            FOneItem.Fbeasongdate   = rsACADEMYget("beasongdate")
			FOneItem.Fmileage		= rsACADEMYget("mileage")

            FOneItem.FImageList     = imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
            FOneItem.FImageSmall    = imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")

            FOneItem.Fjupsustartday		= rsACADEMYget("regstartdate")
            FOneItem.Fjupsuendday		= rsACADEMYget("regenddate")

            FOneItem.Flecturestartday	= rsACADEMYget("lecstartdate")
            FOneItem.Flectureendday		= rsACADEMYget("lecenddate")

            FOneItem.Flimitmaxno		= rsACADEMYget("limit_count")
            FOneItem.Flimitminno		= rsACADEMYget("min_count")
            FOneItem.Flimitsoldno		= rsACADEMYget("limit_sold")
            FOneItem.Flimitwaitno		= rsACADEMYget("wait_count")
            FOneItem.Fitemno			= rsACADEMYget("itemno")
            FOneItem.FWeClassYN			= CHKIIF(isNull(rsACADEMYget("weClassYN")),"",rsACADEMYget("weClassYN"))

        end if
        rsACADEMYget.close
        
        If FOneItem.FWeClassYN = "Y" Then
        	sqlStr = "SELECT * FROM [db_academy].[dbo].[tbl_academy_order_weclass] WHERE orderserial = '" + CStr(FRectOrderSerial) + "'"
	        rsACADEMYget.Open sqlStr,dbACADEMYget,1
	        if  not rsACADEMYget.EOF  then
				FOneItem.FWantStudyName		= rsACADEMYget("wantstudyName")
				FOneItem.FWantStudyYear		= rsACADEMYget("wantstudyYear")
				FOneItem.FWantStudyMonth	= rsACADEMYget("wantstudyMonth")
				FOneItem.FWantStudyDay		= rsACADEMYget("wantstudyDay")
				FOneItem.FWantStudyAmPm		= rsACADEMYget("wantstudyAmPm")
				FOneItem.FWantStudyHour		= rsACADEMYget("wantstudyHour")
				FOneItem.FWantStudyMin		= rsACADEMYget("wantstudyMin")
				FOneItem.FWantStudyPlace	= rsACADEMYget("wantstudyPlace")
				FOneItem.FWantStudyWho		= rsACADEMYget("wantstudyWho")
	        end if
	        rsACADEMYget.close
        End IF
    end Sub


    public Sub GetRequestLectureMasterList()
     	dim sqlStr,i, AddSQL

     	sqlStr = " select count(m.idx) as cnt "
     	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
     	sqlStr = sqlStr + " where 1 = 1 "

		AddSQL =""

     	if (FRectOrderSerial<>"") then
            AddSQL = AddSQL + " and m.orderserial='" + FRectOrderSerial + "' "
        end if

        if (FRectRegStart<>"") then
            AddSQL = AddSQL + " and m.regdate >='" + CStr(FRectRegStart) + "' "
        end if

        if (FRectRegEnd<>"") then
            AddSQL = AddSQL + " and m.regdate <'" + CStr(FRectRegEnd) + "' "
        end if

        if (FRectUserID<>"") then
            AddSQL = AddSQL + " and m.userid='" + FRectUserID + "' "
        end if

        if (FRectBuyname<>"") then
            AddSQL = AddSQL + " and m.buyname like '" + FRectBuyname + "%' "
        end if

        if (FRectReqName<>"") then
            AddSQL = AddSQL + " and m.reqname like '" + FRectReqName + "%' "
        end if

        if (FRectIpkumName<>"") then
            AddSQL = AddSQL + " and m.accountname like '" + FRectIpkumName + "%' "
        end if

        if (FRectSubTotalPrice<>"") then
            AddSQL = AddSQL + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + " "
        end if

        if (FRectBuyHp<>"") then
            AddSQL = AddSQL + " and m.buyhp='" + FRectBuyHp + "' "
        end if

        if (FRectReqHp<>"") then
            AddSQL = AddSQL + " and m.reqhp='" + FRectReqHp + "' "
        end if

        if (FRectBuyPhone<>"") then
            AddSQL = AddSQL + " and m.buyphone='" + FRectBuyPhone + "' "
        end if

        if (FRectReqPhone<>"") then
            AddSQL = AddSQL + " and m.reqphone='" + FRectReqPhone + "' "
        end if

        if (FRectSiteName<>"") then
            AddSQL = AddSQL + " and m.sitename='" + FRectSiteName + "' "
        end if


		''검색조건이 없을경우 FTotalCount=100 최근 100건
		if AddSQL="" then
			FTotalCount = 100
		else
     		rsACADEMYget.Open sqlStr + AddSQL, dbACADEMYget, 1
        		FTotalCount = rsACADEMYget("cnt")
        	rsACADEMYget.Close
        end if

        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
        sqlStr = sqlStr + "     m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv, m.totalitemno, m.totalmileage, m.totalsum, m.discountrate, m.discountprice, m.cancelitemno "
        sqlStr = sqlStr + "     , m.cancelprice, m.subtotalitemno, m.subtotalprice, m.ipkumdiv, m.ipkumdate, m.regdate, m.beadaldate, m.baljudate, m.canceldate, m.cancelyn, m.buyname, m.buyphone "
        sqlStr = sqlStr + "     , m.buyhp, m.buyemail, m.reqname "
        sqlStr = sqlStr + "     , m.reqzipcode, m.reqzipaddr, m.reqaddress, m.reqphone, m.reqhp, m.reqemail, m.songjangdiv, m.deliverno, m.sitename, m.paygatetid, m.resultmsg, m.rduserid, m.milelogid "
        sqlStr = sqlStr + "     , m.miletotalprice, m.jungsanflag, m.authcode, m.rdsite, m.tencardspend, m.reqdate, m.reqtime, m.cardribbon, m.fromname, m.cashreceiptreq "
        sqlStr = sqlStr + "     , m.inireceipttid, m.referip, m.userlevel, m.linkorderserial, m.spendmembership, m.sentenceidx, m.reguserid, m.oldorderserial, m.goodsnames "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by m.idx desc "

        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1

        FtotalPage =  CInt(FTotalCount\FPageSize)
        if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
            FtotalPage = FtotalPage +1
        end if
        FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        redim preserve FItemList(FResultCount)

        if  not rsACADEMYget.EOF  then
            i = 0
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.eof
                set FItemList(i) = new CRequestLectureMasterItem

                FItemList(i).Fidx              = rsACADEMYget("idx")
                FItemList(i).Forderserial      = rsACADEMYget("orderserial")
                FItemList(i).Fjumundiv         = Trim(rsACADEMYget("jumundiv"))
                FItemList(i).Fuserid           = rsACADEMYget("userid")
                FItemList(i).Faccountname      = db2html(rsACADEMYget("accountname"))
                FItemList(i).Faccountdiv       = Trim(rsACADEMYget("accountdiv"))
                FItemList(i).Ftotalitemno      = rsACADEMYget("totalitemno")
                FItemList(i).Ftotalmileage     = rsACADEMYget("totalmileage")
                FItemList(i).Ftotalsum         = rsACADEMYget("totalsum")
                FItemList(i).Fdiscountrate     = rsACADEMYget("discountrate")
                FItemList(i).Fdiscountprice    = rsACADEMYget("discountprice")
                FItemList(i).Fcancelitemno     = rsACADEMYget("cancelitemno")
                FItemList(i).Fcancelprice      = rsACADEMYget("cancelprice")
                FItemList(i).Fsubtotalitemno   = rsACADEMYget("subtotalitemno")
                FItemList(i).Fsubtotalprice    = rsACADEMYget("subtotalprice")
                FItemList(i).Fipkumdiv         = rsACADEMYget("ipkumdiv")
                FItemList(i).Fipkumdate        = rsACADEMYget("ipkumdate")
                FItemList(i).Fregdate          = rsACADEMYget("regdate")
                FItemList(i).Fbeadaldate       = rsACADEMYget("beadaldate")
                FItemList(i).Fbaljudate        = rsACADEMYget("baljudate")
                FItemList(i).Fcanceldate       = rsACADEMYget("canceldate")
                FItemList(i).Fcancelyn         = UCase(rsACADEMYget("cancelyn"))
                FItemList(i).Fbuyname          = db2html(rsACADEMYget("buyname"))
                FItemList(i).Fbuyphone         = rsACADEMYget("buyphone")
                FItemList(i).Fbuyhp            = rsACADEMYget("buyhp")
                FItemList(i).Fbuyemail         = db2html(rsACADEMYget("buyemail"))
                FItemList(i).Freqname          = db2html(rsACADEMYget("reqname"))
                FItemList(i).Freqzipcode       = rsACADEMYget("reqzipcode")
                FItemList(i).Freqzipaddr       = db2html(rsACADEMYget("reqzipaddr"))
                FItemList(i).Freqaddress       = db2html(rsACADEMYget("reqaddress"))
                FItemList(i).Freqphone         = rsACADEMYget("reqphone")
                FItemList(i).Freqhp            = rsACADEMYget("reqhp")
                FItemList(i).Freqemail         = db2html(rsACADEMYget("reqemail"))
                'FItemList(i).Fcomment          = db2html(rsACADEMYget("comment"))
                FItemList(i).Fsitename         = rsACADEMYget("sitename")
                FItemList(i).Fpaygatetid       = rsACADEMYget("paygatetid")
                FItemList(i).Fresultmsg        = rsACADEMYget("resultmsg")
                FItemList(i).Frduserid         = rsACADEMYget("rduserid")
                FItemList(i).Fmilelogid        = rsACADEMYget("milelogid")
                FItemList(i).Fmiletotalprice   = rsACADEMYget("miletotalprice")
                FItemList(i).Fjungsanflag      = rsACADEMYget("jungsanflag")
                FItemList(i).Fauthcode         = rsACADEMYget("authcode")
                FItemList(i).Frdsite           = rsACADEMYget("rdsite")
                FItemList(i).Ftencardspend     = rsACADEMYget("tencardspend")
                'FItemList(i).Fbeasongmemo      = db2html(rsACADEMYget("beasongmemo"))
                FItemList(i).Freqdate          = rsACADEMYget("reqdate")
                FItemList(i).Freqtime          = rsACADEMYget("reqtime")
                FItemList(i).Fcardribbon       = rsACADEMYget("cardribbon")
                'FItemList(i).Fmessage          = db2html(rsACADEMYget("message"))
                FItemList(i).Ffromname         = db2html(rsACADEMYget("fromname"))
                FItemList(i).Fcashreceiptreq   = rsACADEMYget("cashreceiptreq")
                FItemList(i).Finireceipttid    = rsACADEMYget("inireceipttid")
                FItemList(i).Freferip          = rsACADEMYget("referip")
                FItemList(i).Fuserlevel        = rsACADEMYget("userlevel")
                FItemList(i).Flinkorderserial  = rsACADEMYget("linkorderserial")
                FItemList(i).Fspendmembership  = rsACADEMYget("spendmembership")
                FItemList(i).Fsentenceidx      = rsACADEMYget("sentenceidx")
                FItemList(i).Freguserid        = rsACADEMYget("reguserid")
                FItemList(i).Foldorderserial   = rsACADEMYget("oldorderserial")

				FItemList(i).Fgoodsnames		= db2html(rsACADEMYget("goodsnames"))

                rsACADEMYget.MoveNext
                i = i + 1
            loop
        end if
        rsACADEMYget.close

    end Sub



    'TODO : 디테일에는 하나의 강좌만 있다고 가정한다.
    '' 심플하게 수정 해야함 마스터만 사용....
    public Sub GetRequestLectureMasterList_OLD()
        dim sqlStr,i

        sqlStr = " select distinct m.idx "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "

        'sqlStr = " select count(*) as cnt from [db_academy].[dbo].tbl_academy_order_master m "
        'sqlStr = sqlStr + " where 1 = 1 "
        ''sqlStr = sqlStr + " and m.ipkumdiv>1"
        ''sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.jumundiv = '8' "

        if (FRectOrderSerial<>"") then
            sqlStr = sqlStr + " and m.orderserial='" + FRectOrderSerial + "' "
        end if

        if (FRectRegStart<>"") then
            sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "' "
        end if

        if (FRectRegEnd<>"") then
            sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "' "
        end if

        if (FRectUserID<>"") then
            sqlStr = sqlStr + " and m.userid='" + FRectUserID + "' "
        end if

        if (FRectBuyname<>"") then
            sqlStr = sqlStr + " and m.buyname like '" + FRectBuyname + "%' "
        end if

        if (FRectReqName<>"") then
            sqlStr = sqlStr + " and m.reqname like '" + FRectReqName + "%' "
        end if

        if (FRectIpkumName<>"") then
            sqlStr = sqlStr + " and m.accountname like '" + FRectIpkumName + "%' "
        end if

        if (FRectSubTotalPrice<>"") then
            sqlStr = sqlStr + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + " "
        end if

        if (FRectBuyHp<>"") then
            sqlStr = sqlStr + " and m.buyhp='" + FRectBuyHp + "' "
        end if

        if (FRectReqHp<>"") then
            sqlStr = sqlStr + " and m.reqhp='" + FRectReqHp + "' "
        end if

        if (FRectBuyPhone<>"") then
            sqlStr = sqlStr + " and m.buyphone='" + FRectBuyPhone + "' "
        end if

        if (FRectReqPhone<>"") then
            sqlStr = sqlStr + " and m.reqphone='" + FRectReqPhone + "' "
        end if

        if (FRectSiteName<>"") then
            sqlStr = sqlStr + " and m.sitename='" + FRectSiteName + "' "
        end if

        if (FRectItemID<>"") then
            sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemID) + " "
        end if

        sqlStr = " select count(T.idx) as cnt from (" + sqlStr + ") T "





        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		'FTotalCount = 100

        sqlStr = " select distinct top " + CStr(FPageSize*FCurrPage) + " "
        sqlStr = sqlStr + "     m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv, m.totalitemno, m.totalmileage, m.totalsum, m.discountrate, m.discountprice, m.cancelitemno "
        sqlStr = sqlStr + "     , m.cancelprice, m.subtotalitemno, m.subtotalprice, m.ipkumdiv, m.ipkumdate, m.regdate, m.beadaldate, m.baljudate, m.canceldate, m.cancelyn, m.buyname, m.buyphone "
        sqlStr = sqlStr + "     , m.buyhp, m.buyemail, m.reqname "
        sqlStr = sqlStr + "     , m.reqzipcode, m.reqzipaddr, m.reqaddress, m.reqphone, m.reqhp, m.reqemail, m.songjangdiv, m.deliverno, m.sitename, m.paygatetid, m.resultmsg, m.rduserid, m.milelogid "
        sqlStr = sqlStr + "     , m.miletotalprice, m.jungsanflag, m.authcode, m.rdsite, m.tencardspend, m.reqdate, m.reqtime, m.cardribbon, m.fromname, m.cashreceiptreq "
        sqlStr = sqlStr + "     , m.inireceipttid, m.referip, m.userlevel, m.linkorderserial, m.spendmembership, m.sentenceidx, m.reguserid, m.oldorderserial, m.goodsnames "
        sqlStr = sqlStr + "     , d.itemid "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "



        'sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        'sqlStr = sqlStr + " m.* from [db_academy].[dbo].tbl_academy_order_master m "
        'sqlStr = sqlStr + " where 1 = 1 "
		''sqlStr = sqlStr + " and m.ipkumdiv>1"
        ''sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.jumundiv = '8' "

        if (FRectOrderSerial<>"") then
            sqlStr = sqlStr + " and m.orderserial='" + FRectOrderSerial + "' "
        end if

        if (FRectRegStart<>"") then
            sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "' "
        end if

        if (FRectRegEnd<>"") then
            sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "' "
        end if

        if (FRectUserID<>"") then
            sqlStr = sqlStr + " and m.userid='" + FRectUserID + "' "
        end if

        if (FRectBuyname<>"") then
            sqlStr = sqlStr + " and m.buyname like '" + FRectBuyname + "%' "
        end if

        if (FRectReqName<>"") then
            sqlStr = sqlStr + " and m.reqname like '" + FRectReqName + "%' "
        end if

        if (FRectIpkumName<>"") then
            sqlStr = sqlStr + " and m.accountname like '" + FRectIpkumName + "%' "
        end if

        if (FRectSubTotalPrice<>"") then
            sqlStr = sqlStr + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + " "
        end if

        if (FRectBuyHp<>"") then
            sqlStr = sqlStr + " and m.buyhp='" + FRectBuyHp + "' "
        end if

        if (FRectReqHp<>"") then
            sqlStr = sqlStr + " and m.reqhp='" + FRectReqHp + "' "
        end if

        if (FRectBuyPhone<>"") then
            sqlStr = sqlStr + " and m.buyphone='" + FRectBuyPhone + "' "
        end if

        if (FRectReqPhone<>"") then
            sqlStr = sqlStr + " and m.reqphone='" + FRectReqPhone + "' "
        end if

        if (FRectSiteName<>"") then
            sqlStr = sqlStr + " and m.sitename='" + FRectSiteName + "' "
        end if

        if (FRectItemID<>"") then
            sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemID) + " "
        end if

        sqlStr = sqlStr + " order by m.idx desc "

        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        'response.write sqlStr
        FtotalPage =  CInt(FTotalCount\FPageSize)
        if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
            FtotalPage = FtotalPage +1
        end if
        FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        redim preserve FItemList(FResultCount)

        if  not rsACADEMYget.EOF  then
            i = 0
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.eof
                set FItemList(i) = new CRequestLectureMasterItem

                FItemList(i).Fidx              = rsACADEMYget("idx")
                FItemList(i).Forderserial      = rsACADEMYget("orderserial")
                FItemList(i).Fjumundiv         = Trim(rsACADEMYget("jumundiv"))
                FItemList(i).Fuserid           = rsACADEMYget("userid")
                FItemList(i).Faccountname      = db2html(rsACADEMYget("accountname"))
                FItemList(i).Faccountdiv       = Trim(rsACADEMYget("accountdiv"))
                FItemList(i).Ftotalitemno      = rsACADEMYget("totalitemno")
                FItemList(i).Ftotalmileage     = rsACADEMYget("totalmileage")
                FItemList(i).Ftotalsum         = rsACADEMYget("totalsum")
                FItemList(i).Fdiscountrate     = rsACADEMYget("discountrate")
                FItemList(i).Fdiscountprice    = rsACADEMYget("discountprice")
                FItemList(i).Fcancelitemno     = rsACADEMYget("cancelitemno")
                FItemList(i).Fcancelprice      = rsACADEMYget("cancelprice")
                FItemList(i).Fsubtotalitemno   = rsACADEMYget("subtotalitemno")
                FItemList(i).Fsubtotalprice    = rsACADEMYget("subtotalprice")
                FItemList(i).Fipkumdiv         = rsACADEMYget("ipkumdiv")
                FItemList(i).Fipkumdate        = rsACADEMYget("ipkumdate")
                FItemList(i).Fregdate          = rsACADEMYget("regdate")
                FItemList(i).Fbeadaldate       = rsACADEMYget("beadaldate")
                FItemList(i).Fbaljudate        = rsACADEMYget("baljudate")
                FItemList(i).Fcanceldate       = rsACADEMYget("canceldate")
                FItemList(i).Fcancelyn         = UCase(rsACADEMYget("cancelyn"))
                FItemList(i).Fbuyname          = db2html(rsACADEMYget("buyname"))
                FItemList(i).Fbuyphone         = rsACADEMYget("buyphone")
                FItemList(i).Fbuyhp            = rsACADEMYget("buyhp")
                FItemList(i).Fbuyemail         = db2html(rsACADEMYget("buyemail"))
                FItemList(i).Freqname          = db2html(rsACADEMYget("reqname"))
                FItemList(i).Freqzipcode       = rsACADEMYget("reqzipcode")
                FItemList(i).Freqzipaddr       = db2html(rsACADEMYget("reqzipaddr"))
                FItemList(i).Freqaddress       = db2html(rsACADEMYget("reqaddress"))
                FItemList(i).Freqphone         = rsACADEMYget("reqphone")
                FItemList(i).Freqhp            = rsACADEMYget("reqhp")
                FItemList(i).Freqemail         = db2html(rsACADEMYget("reqemail"))
                'FItemList(i).Fcomment          = db2html(rsACADEMYget("comment"))
                FItemList(i).Fsitename         = rsACADEMYget("sitename")
                FItemList(i).Fpaygatetid       = rsACADEMYget("paygatetid")
                FItemList(i).Fresultmsg        = rsACADEMYget("resultmsg")
                FItemList(i).Frduserid         = rsACADEMYget("rduserid")
                FItemList(i).Fmilelogid        = rsACADEMYget("milelogid")
                FItemList(i).Fmiletotalprice   = rsACADEMYget("miletotalprice")
                FItemList(i).Fjungsanflag      = rsACADEMYget("jungsanflag")
                FItemList(i).Fauthcode         = rsACADEMYget("authcode")
                FItemList(i).Frdsite           = rsACADEMYget("rdsite")
                FItemList(i).Ftencardspend     = rsACADEMYget("tencardspend")
                'FItemList(i).Fbeasongmemo      = db2html(rsACADEMYget("beasongmemo"))
                FItemList(i).Freqdate          = rsACADEMYget("reqdate")
                FItemList(i).Freqtime          = rsACADEMYget("reqtime")
                FItemList(i).Fcardribbon       = rsACADEMYget("cardribbon")
                'FItemList(i).Fmessage          = db2html(rsACADEMYget("message"))
                FItemList(i).Ffromname         = db2html(rsACADEMYget("fromname"))
                FItemList(i).Fcashreceiptreq   = rsACADEMYget("cashreceiptreq")
                FItemList(i).Finireceipttid    = rsACADEMYget("inireceipttid")
                FItemList(i).Freferip          = rsACADEMYget("referip")
                FItemList(i).Fuserlevel        = rsACADEMYget("userlevel")
                FItemList(i).Flinkorderserial  = rsACADEMYget("linkorderserial")
                FItemList(i).Fspendmembership  = rsACADEMYget("spendmembership")
                FItemList(i).Fsentenceidx      = rsACADEMYget("sentenceidx")
                FItemList(i).Freguserid        = rsACADEMYget("reguserid")
                FItemList(i).Foldorderserial   = rsACADEMYget("oldorderserial")

                rsACADEMYget.MoveNext
                i = i + 1
            loop
        end if
        rsACADEMYget.close

        '강좌정보
        dim masteridxlist, j
        for i = 0 to (FResultCount - 1)
                if (i = 0) then
                        masteridxlist = CStr(FItemList(i).Fidx)
                else
                        masteridxlist = masteridxlist + "','" + CStr(FItemList(i).Fidx)
                end if
        next

		masteridxlist = "'" + masteridxlist + "'"

        'TODO : 디테일정보에는 하나의 강좌정보만 있다고 가정한다.
        sqlStr = " select distinct d.orderserial, d.itemcost, d.buycash, d.itemname, d.itemid, d.itemoption, d.makerid, d.itemoptionname, d.beasongdate, l.listimg, l.smallimg, l.lecturer_name "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d, [db_academy].[dbo].tbl_lec_item l "
        sqlStr = sqlStr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
        sqlStr = sqlStr + " and d.itemid = l.idx "
        'response.write sqlStr
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        if  not rsACADEMYget.EOF  then
            do until rsACADEMYget.eof
                for j = 0 to (FResultCount - 1)
                    if (FItemList(j).Forderserial = rsACADEMYget("orderserial")) then
                        FItemList(j).Fitemcost      = rsACADEMYget("itemcost")
                        FItemList(j).Fbuycash       = rsACADEMYget("buycash")
                        FItemList(j).Fitemname      = db2html(rsACADEMYget("itemname"))
                        FItemList(j).Fitemid        = rsACADEMYget("itemid")
                        FItemList(j).Fitemoption    = rsACADEMYget("itemoption")
                        FItemList(j).Fmakerid       = rsACADEMYget("makerid")
                        FItemList(j).Fmakername     = db2html(rsACADEMYget("lecturer_name"))
                        FItemList(j).Fitemoptionname    = db2html(rsACADEMYget("itemoptionname"))
                        FItemList(j).Fbeasongdate   = rsACADEMYget("beasongdate")

                        FItemList(j).FImageList     = imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
                        FItemList(j).FImageSmall    = imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")
                        exit for
                    end if
                next
                rsACADEMYget.moveNext
            loop
        end if
        rsACADEMYget.close
    end Sub

    public Sub CRequestLectureDetailList()
        dim sqlStr
        dim i

		'itemcost : 합계 / reducedprice : 할인적용금액 / matcostAdded : 인당 재료비
        sqlStr = "select d.detailidx, d.orderserial,d.oitemdiv, d.itemid,d.itemoption,d.itemno,d.itemcost,d.buycash, d.reducedprice, d.matcostAdded, d.bonuscouponidx, d.leccouponidx as itemcouponidx, "
        sqlStr = sqlStr + " d.mileage,d.cancelyn,"
        sqlStr = sqlStr + " d.itemname, d.makerid, i.listimg, i.smallimg ,"
        sqlStr = sqlStr + " d.itemoptionname , d.currstate, d.upcheconfirmdate, d.songjangdiv, "
        sqlStr = sqlStr + " d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem, d.entryname, d.entryhp, d.requiredetail, s.divname as songjangdivname, s.findurl  "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item i, "
        sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
        sqlStr = sqlStr + "     left join [db_academy].[dbo].tbl_songjang_div s on d.songjangdiv=s.divcd"
        sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
        sqlStr = sqlStr + " and d.itemid=i.idx"
        sqlStr = sqlStr + " order by d.detailidx "
        'response.write sqlStr
        rsACADEMYget.Open sqlStr,dbACADEMYget,1

        FTotalCount = rsACADEMYget.RecordCount
        FResultCount = FTotalCount

        redim preserve FItemList(FResultCount)

        i=0
        do until rsACADEMYget.eof
                set FItemList(i) = new CRequestLectureDetailItem

                FItemList(i).Forderserial       = CStr(FRectOrderSerial)
                FItemList(i).Fdetailidx         = rsACADEMYget("detailidx")

                FItemList(i).Fentryname         = db2html(rsACADEMYget("entryname"))
                FItemList(i).Fentryhp           = rsACADEMYget("entryhp")
                FItemList(i).Fcancelyn          = UCase(rsACADEMYget("cancelyn"))

                FItemList(i).Fitemoptionname    = db2html(rsACADEMYget("itemoptionname"))

				FItemList(i).Fmakerid      		= rsACADEMYget("makerid")
				FItemList(i).Foitemdiv      	= rsACADEMYget("oitemdiv")
				FItemList(i).Fitemid      		= rsACADEMYget("itemid")
				FItemList(i).Fitemoption  		= rsACADEMYget("itemoption")
				FItemList(i).Fitemno      		= rsACADEMYget("itemno")
				FItemList(i).Fitemcost    		= rsACADEMYget("itemcost")
				FItemList(i).Fbuycash     		= rsACADEMYget("buycash")
				FItemList(i).Fmileage     		= rsACADEMYget("mileage")
				'FItemList(i).Fcosttotal   		= rsACADEMYget("costtotal")
				'FItemList(i).Forderdate   		= rsACADEMYget("orderdate")

				FItemList(i).FItemName    		= db2html(rsACADEMYget("itemname"))
				'FItemList(i).FImageList    		= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
				'FItemList(i).FImageSmall    	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("smallimage")

	            FItemList(i).FImageList     	= imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
	            FItemList(i).FImageSmall    	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")

				FItemList(i).Fcurrstate     	= rsACADEMYget("currstate")
				FItemList(i).Fsongjangdiv   	= rsACADEMYget("songjangdiv")
				FItemList(i).Fsongjangno    	= rsACADEMYget("songjangno")
				FItemList(i).Fupcheconfirmdate 	= rsACADEMYget("upcheconfirmdate")
				FItemList(i).Fbeasongdate   	= rsACADEMYget("beasongdate")
				FItemList(i).Fisupchebeasong	= rsACADEMYget("isupchebeasong")
				FItemList(i).Fissailitem    	= rsACADEMYget("issailitem")

				FItemList(i).Frequiredetail    	= rsACADEMYget("requiredetail")

				FItemList(i).Fsongjangdivname  	= rsACADEMYget("songjangdivname")
				FItemList(i).Ffindurl  			= rsACADEMYget("findurl")

				FItemList(i).Freducedprice  	= rsACADEMYget("reducedprice")
				FItemList(i).FmatcostAdded  	= rsACADEMYget("matcostAdded")

	            FItemList(i).Fbonuscouponidx    = rsACADEMYget("bonuscouponidx")
	            FItemList(i).Fitemcouponidx     = rsACADEMYget("itemcouponidx")

				if IsNull(FItemList(i).Freducedprice) = True then
					FItemList(i).Freducedprice  	= rsACADEMYget("itemcost")
				end if

                rsACADEMYget.movenext
                i=i+1
        loop
        rsACADEMYget.close
    end Sub

    public Sub CRequestDIYItemDetailList()
        dim sqlStr
        dim i

		'itemcost : 합계 / reducedprice : 할인적용금액 / matcostAdded : 인당 재료비
        sqlStr = "select d.detailidx, d.orderserial,d.oitemdiv, d.itemid,d.itemoption,d.itemno,d.itemcost,d.buycash, d.reducedprice, d.matcostAdded, d.bonuscouponidx, d.leccouponidx as itemcouponidx, "
        sqlStr = sqlStr + " d.mileage,d.cancelyn,"
        sqlStr = sqlStr + " d.itemname, d.makerid, i.listimage as listimg, i.smallimage as smallimg ,"
        sqlStr = sqlStr + " d.itemoptionname , d.currstate, d.upcheconfirmdate, d.songjangdiv, "
        sqlStr = sqlStr + " d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem, d.entryname, d.entryhp, d.requiredetail, s.divname as songjangdivname, s.findurl  "

        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d "
        sqlStr = sqlStr + "  	left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid "
        sqlStr = sqlStr + "     left join [db_academy].[dbo].tbl_songjang_div s on d.songjangdiv=s.divcd"
        sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
        sqlStr = sqlStr + " order by d.detailidx "
        'response.write sqlStr
        rsACADEMYget.Open sqlStr,dbACADEMYget,1

        FTotalCount = rsACADEMYget.RecordCount
        FResultCount = FTotalCount

        redim preserve FItemList(FResultCount)

        i=0
        do until rsACADEMYget.eof
                set FItemList(i) = new CRequestLectureDetailItem

                FItemList(i).Forderserial       = CStr(FRectOrderSerial)
                FItemList(i).Fdetailidx         = rsACADEMYget("detailidx")

                FItemList(i).Fentryname         = db2html(rsACADEMYget("entryname"))
                FItemList(i).Fentryhp           = rsACADEMYget("entryhp")
                FItemList(i).Fcancelyn          = UCase(rsACADEMYget("cancelyn"))

                FItemList(i).Fitemoptionname    = db2html(rsACADEMYget("itemoptionname"))

				FItemList(i).Fmakerid      		= rsACADEMYget("makerid")
				FItemList(i).Foitemdiv      	= rsACADEMYget("oitemdiv")
				FItemList(i).Fitemid      		= rsACADEMYget("itemid")
				FItemList(i).Fitemoption  		= rsACADEMYget("itemoption")
				FItemList(i).Fitemno      		= rsACADEMYget("itemno")
				FItemList(i).Fitemcost    		= rsACADEMYget("itemcost")
				FItemList(i).Fbuycash     		= rsACADEMYget("buycash")
				FItemList(i).Fmileage     		= rsACADEMYget("mileage")
				'FItemList(i).Fcosttotal   		= rsACADEMYget("costtotal")
				'FItemList(i).Forderdate   		= rsACADEMYget("orderdate")

				FItemList(i).FItemName    		= db2html(rsACADEMYget("itemname"))
				'FItemList(i).FImageList    		= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
				'FItemList(i).FImageSmall    	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("smallimage")

				'강좌는 디테일 이미지가 없다. DIY 상품만 이미지가 있다.

				FItemList(i).FImageList			= imgFingers & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
				FItemList(i).FImageSmall		= imgFingers & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")

'response.write FItemList(i).FImageSmall & "aaaaaaaaaa" & rsACADEMYget("itemid")

	            'FItemList(i).FImageList     	= imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
	            'FItemList(i).FImageSmall    	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")

				FItemList(i).Fcurrstate     	= rsACADEMYget("currstate")
				FItemList(i).Fsongjangdiv   	= rsACADEMYget("songjangdiv")
				FItemList(i).Fsongjangno    	= rsACADEMYget("songjangno")
				FItemList(i).Fupcheconfirmdate 	= rsACADEMYget("upcheconfirmdate")
				FItemList(i).Fbeasongdate   	= rsACADEMYget("beasongdate")
				FItemList(i).Fisupchebeasong	= rsACADEMYget("isupchebeasong")
				FItemList(i).Fissailitem    	= rsACADEMYget("issailitem")

				FItemList(i).Frequiredetail    	= rsACADEMYget("requiredetail")

				FItemList(i).Fsongjangdivname  	= rsACADEMYget("songjangdivname")
				FItemList(i).Ffindurl  			= rsACADEMYget("findurl")

				FItemList(i).Freducedprice  	= rsACADEMYget("reducedprice")
				FItemList(i).FmatcostAdded  	= rsACADEMYget("matcostAdded")

	            FItemList(i).Fbonuscouponidx    = rsACADEMYget("bonuscouponidx")
	            FItemList(i).Fitemcouponidx     = rsACADEMYget("itemcouponidx")

				if IsNull(FItemList(i).Freducedprice) = True then
					FItemList(i).Freducedprice  	= rsACADEMYget("itemcost")
				end if

                rsACADEMYget.movenext
                i=i+1
        loop
        rsACADEMYget.close
    end Sub

    Private Sub Class_Initialize()
        redim  FItemList(0)
        FCurrPage =1
        FPageSize = 100
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

function LectureIpkumDivColor(byval k)
	if k="0" then
		LectureIpkumDivColor="#AAAA00"
	elseif k="1" then
		LectureIpkumDivColor"#AAAA00"
	elseif k="2" then
		LectureIpkumDivColor"#AAAAAA"
	elseif k="3" then
		LectureIpkumDivColor="#AAAAAA"
	elseif k="4" then
		LectureIpkumDivColor="#0000FF"
	elseif k="5" then
		LectureIpkumDivColor="#6C6C6C"
	elseif k="6" then
		LectureIpkumDivColor="#33AAAA"
	elseif k="7" then
		LectureIpkumDivColor="#FF0000"
	elseif k="8" then
		LectureIpkumDivColor="FF0000"
	else
		LectureIpkumDivColor="#FFFFFF"
	end if
end function

function LectureIpkumDivName(byval k)
	if k="0" then
		LectureIpkumDivName="주문실패"
	elseif k="1" then
		LectureIpkumDivName="주문실패"
	elseif k="2" then
		LectureIpkumDivName="입금대기"
	elseif k="3" then
		LectureIpkumDivName="입금대기"
	elseif k="4" then
		LectureIpkumDivName="결제완료"
	elseif k="5" then
		LectureIpkumDivName="강좌준비"
	elseif k="6" then
		LectureIpkumDivName="강좌확정"
	elseif k="7" then
		LectureIpkumDivName="강좌확정"
	elseif k="8" then
		LectureIpkumDivName="강좌확정"
	end if
end function

Function fnWeClassStudyWho(v)
	SELECT CASE v
		Case "1" : fnWeClassStudyWho = "기업"
		Case "2" : fnWeClassStudyWho = "동호회"
		Case "3" : fnWeClassStudyWho = "학생"
		Case "0" : fnWeClassStudyWho = "기타"
	End SELECT
End Function

%>