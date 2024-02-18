<%
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
			IpkumDivColor="#00FF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
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

end class

class CJumunDetailItem
	public Forderserial
	public Fdetailidx

	public FMakerid
	public Fitemid
	public Fitemoption
	public Fitemno
	public Fitemcost
	public Fbuycash
	public Fitemvat
	public Fmileage
	''public Fcosttotal
	''public Forderdate
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
		dim i
		for i=0 to FDetailCount-1
			if FJumunDetailList(i).FItemID=0 then
				BeasongPay = FJumunDetailList(i).Fitemcost
				Exit For
			end if
		next
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

	public FDtlItemName
	public FDtlItemNo
	public FDtlItemOption
	public FDtlItemOptionName
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Freqdate
	public Freqtime

    public FDlvcountryCode

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

    ''해외배송인지여부
	public function IsForeignDeliver()
        IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR")
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
				IpkumDivColor="#33AA33"
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
	public FRectUserID
	public FRectckdate
	public FRectBuyname
	public FRectReqName
	public FRectIpkumName
	public FRectSubTotalPrice

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

	public FRectOldJumun

    public FRectIpkumDiv
    public FRectMinusOrderInclude

    public FRectIsUpcheBeasong

    public FRectIsFlower
    public FRectIsMinus
    public FRectIsForeign
    public FRectIsMilitary

	Private Sub Class_Initialize()
'		redim preserve FMasterItemList(0)
		redim FMasterItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
		sqlStr = "select count(itemid) as cnt from [db_academy].[dbo].tbl_diy_item"
		sqlStr = sqlStr + " where makerid='" + session("ssBctID") + "'"
		sqlStr = sqlStr + " and deliverytype<>'1'"
		sqlStr = sqlStr + " and deliverytype<>'4'"
		sqlStr = sqlStr + " and sellyn<>'N'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		IsAllTenBeasong = (rsACADEMYget("cnt")<1)
		rsACADEMYget.close
	end function

	public sub SearchTargetItemJumunList()
		dim sqlStr
		dim i

		sqlStr = "select top 1000 m.orderserial, m.buyname, m.reqname, m.ipkumdiv, m.jumundiv,"
		sqlStr = sqlStr + " convert(varchar(10),m.regdate,21) as regdate, convert(varchar(10),m.ipkumdate,21) as ipkumdate, m.comment , d.itemname, d.itemno,"
		sqlStr = sqlStr + " d.itemoption, d.itemoptionname"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
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
		sqlStr = sqlStr + " and m.sitename='diyitem'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		if FRectItemid<>"" then
		sqlStr = sqlStr + " and d.itemid=" + FRectItemid + ""
		end if
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " order by m.idx"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fregdate		= rsACADEMYget("regdate")
			FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")

			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
			FMasterItemList(i).FDtlItemName = db2Html(rsACADEMYget("itemname"))
			FMasterItemList(i).FDtlItemNo	= rsACADEMYget("itemno")
			FMasterItemList(i).FDtlItemOption = rsACADEMYget("itemoption")
			FMasterItemList(i).FDtlItemOptionName = db2Html(rsACADEMYget("itemoptionname"))
			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close

	end Sub

	public sub SearchOneJumunDetail(byval idx)
		dim sqlStr
		dim i

		sqlStr = "select d.*, convert(varchar(19),d.upcheconfirmdate,21) as cvupcheconfirmdate, convert(varchar(19),d.beasongdate,21) as cvbeasongdate, i.smallimage,i.listimage, i.sellcash as currsellcash, i.buycash as currbuycash"
		sqlStr = sqlStr + "  from [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where d.detailidx=" + CStr(idx)

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		set FJumunDetail = new CJumunDetailItem

		if Not rsACADEMYget.Eof then
			FJumunDetail.Forderserial = rsACADEMYget("orderserial")
			FJumunDetail.Fdetailidx			= rsACADEMYget("idx")
			FJumunDetail.Fmakerid      = rsACADEMYget("makerid")
			FJumunDetail.Fitemid      = rsACADEMYget("itemid")
			FJumunDetail.Fitemoption  = rsACADEMYget("itemoption")
			FJumunDetail.Fitemno      = rsACADEMYget("itemno")
			FJumunDetail.Fitemcost    = rsACADEMYget("itemcost")
			FJumunDetail.Fbuycash     = rsACADEMYget("buycash")
			FJumunDetail.Fmileage     = rsACADEMYget("mileage")
			'FJumunDetail.Fcosttotal   = rsACADEMYget("costtotal")
			'FJumunDetail.Forderdate   = rsACADEMYget("orderdate")
			FJumunDetail.Fcancelyn    = rsACADEMYget("cancelyn")

			FJumunDetail.FItemName    = db2html(rsACADEMYget("itemname"))
			FJumunDetail.FImageList		= imgFingers & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage")
			FJumunDetail.FImageSmall	= imgFingers & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimage")

			FJumunDetail.FItemoptionName = db2html(rsACADEMYget("itemoptionname"))

			FJumunDetail.Fcurrstate     = rsACADEMYget("currstate")
			FJumunDetail.Fsongjangdiv   = rsACADEMYget("songjangdiv")
			FJumunDetail.Fsongjangno    = rsACADEMYget("songjangno")
			FJumunDetail.Fupcheconfirmdate = rsACADEMYget("cvupcheconfirmdate")
			FJumunDetail.Fbeasongdate   = rsACADEMYget("cvbeasongdate")
			FJumunDetail.Fisupchebeasong= rsACADEMYget("isupchebeasong")
			FJumunDetail.Fissailitem    = rsACADEMYget("issailitem")

			FJumunDetail.Frequiredetail = db2html(rsACADEMYget("requiredetail"))

			FJumunDetail.FcurrSellcash	= rsACADEMYget("currsellcash")
			FJumunDetail.FcurrBuycash	= rsACADEMYget("currbuycash")
            FJumunDetail.Foitemdiv      = rsACADEMYget("oitemdiv")

            FJumunDetail.FOmwDiv        = rsACADEMYget("omwdiv")
            FJumunDetail.FODlvType      = rsACADEMYget("odlvtype")
		end if

		rsACADEMYget.close

	end sub

	public sub SearchJumunDetail(byval orderserial)
		dim sqlStr
		dim i

		sqlStr = "select d.detailidx as idx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.mileage,d.cancelyn,"
		sqlStr = sqlStr + " d.itemname, d.makerid, i.listimage as imglist, i.smallimage as imgsmall, d.itemoptionname as codeview, d.currstate, d.songjangdiv, d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem, d.requiredetail  "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_diy_item i, [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		set FJumunDetail = new CJumunDetail
		FJumunDetail.SetDetailCount rsACADEMYget.RecordCount

		i=0
		do until rsACADEMYget.eof
			set FJumunDetail.FJumunDetailList(i) = new CJumunDetailItem

			FJumunDetail.FJumunDetailList(i).Forderserial = CStr(orderserial)
			FJumunDetail.FJumunDetailList(i).Fdetailidx			= rsACADEMYget("idx")
			FJumunDetail.FJumunDetailList(i).Fmakerid      = rsACADEMYget("makerid")
			FJumunDetail.FJumunDetailList(i).Fitemid      = rsACADEMYget("itemid")
			FJumunDetail.FJumunDetailList(i).Fitemoption  = rsACADEMYget("itemoption")
			FJumunDetail.FJumunDetailList(i).Fitemno      = rsACADEMYget("itemno")
			FJumunDetail.FJumunDetailList(i).Fitemcost    = rsACADEMYget("itemcost")
			'FJumunDetail.FJumunDetailList(i).Fitemvat     = rsACADEMYget("itemvat")
			FJumunDetail.FJumunDetailList(i).Fmileage     = rsACADEMYget("mileage")
			''FJumunDetail.FJumunDetailList(i).Fcosttotal   = rsACADEMYget("costtotal")
			''FJumunDetail.FJumunDetailList(i).Forderdate   = rsACADEMYget("orderdate")
			FJumunDetail.FJumunDetailList(i).Fcancelyn    = rsACADEMYget("cancelyn")

			FJumunDetail.FJumunDetailList(i).FItemName    = db2html(rsACADEMYget("itemname"))
			'FJumunDetail.FJumunDetailList(i).FImageList    = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName(i) + "/" + rsACADEMYget("imglist")
			FJumunDetail.FJumunDetailList(i).FImageList		= imgFingers & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("imglist")
			'FJumunDetail.FJumunDetailList(i).FImageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("imgsmall")
			FJumunDetail.FJumunDetailList(i).FImageSmall	= imgFingers & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("imgsmall")

			if IsNull(rsACADEMYget("codeview")) then
				FJumunDetail.FJumunDetailList(i).FItemoptionName = "-"
			else
				FJumunDetail.FJumunDetailList(i).FItemoptionName = db2html(rsACADEMYget("codeview"))
			end if

			FJumunDetail.FJumunDetailList(i).Fcurrstate     = rsACADEMYget("currstate")
			FJumunDetail.FJumunDetailList(i).Fsongjangdiv   = rsACADEMYget("songjangdiv")
			FJumunDetail.FJumunDetailList(i).Fsongjangno    = rsACADEMYget("songjangno")
			FJumunDetail.FJumunDetailList(i).Fbeasongdate   = rsACADEMYget("beasongdate")
			FJumunDetail.FJumunDetailList(i).Fisupchebeasong= rsACADEMYget("isupchebeasong")
			FJumunDetail.FJumunDetailList(i).Fissailitem    = rsACADEMYget("issailitem")

			FJumunDetail.FJumunDetailList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))
			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.close
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
		wheredetail = wheredetail + " and m.sitename='diyitem'"


		''########## 총 갯수 ################''
		sqlStr = "select count(m.orderserial) as cnt"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where orderserial<>''"
		sqlStr = sqlStr + wheredetail


		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		''########## 데이타 ################''
		sqlStr = "select top " + CStr(FPageSize) + " m.orderserial, m.buyname, m.userid, m.accountdiv, m.sitename, "
		sqlStr = sqlStr + " m.totalsum, m.subtotalprice, m.ipkumdiv, m.regdate, "
		sqlStr = sqlStr + " m.discountrate, m.buyname, m.reqname, m.cancelyn"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where orderserial not in ("
		sqlStr = sqlStr + " select top " + CStr((FCurrPage-1)*FPageSize)  + " orderserial from [db_academy].[dbo].tbl_academy_order_master "
		sqlStr = sqlStr + " where orderserial<>''"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by regdate desc"
		sqlStr = sqlStr + ")"

		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by regdate desc"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsACADEMYget("regdate")
			FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
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
		wheredetail = wheredetail + " and m.sitename='diyitem'"

		''########## 총 갯수 ################''
		sqlStr = "select count(m.orderserial) as cnt"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"

		sqlStr = sqlStr + wheredetail

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		''########## 데이타 ################''
		sqlStr = "select top 100 m.orderserial, m.buyname, m.userid, m.accountdiv, m.sitename, "
		sqlStr = sqlStr + " m.totalsum, m.subtotalprice, m.ipkumdiv, m.regdate, "
		sqlStr = sqlStr + " m.discountrate, m.buyname, m.reqname, m.cancelyn"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"

		sqlStr = sqlStr + wheredetail

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsACADEMYget("regdate")
			FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
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
				wheredetail = wheredetail + " and ((d.isupchebeasong='Y') and (d.beasongdate >='" + CStr(FRectRegStart) + "')) "
			elseif FRectDateType="tenbeasongdate" then
				wheredetail = wheredetail + " and ((d.isupchebeasong='N') and (m.beadaldate >='" + CStr(FRectRegStart) + "'))"
			else
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if FRectDateType="ipkumil" then
				wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif FRectDateType="upbeasongdate" then
				wheredetail = wheredetail + " and ((d.isupchebeasong='Y') and (d.beasongdate <'" + CStr(FRectRegEnd) + "')) "
			elseif FRectDateType="tenbeasongdate" then
				wheredetail = wheredetail + " and ((d.isupchebeasong='N') and (m.beadaldate <'" + CStr(FRectRegEnd) + "'))"
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

		''#################################################
		''총 갯수. 총금액
		''#################################################
		sqlStr = "select count(m.orderserial) as cnt, sum(d.buycash*d.itemno) as totalbuycash"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.sitename='diyitem'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1'"             ''2009 변경, 주문접수건도 표시
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail

        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		    FTotalBuyCash = rsACADEMYget("totalbuycash")
		    FTotalCount = rsACADEMYget("cnt")

		    if IsNULL(FTotalBuyCash) then FTotalBuyCash=0
		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " d.orderserial, m.buyname,m.reqname, m.jumundiv, m.userid,"
		sqlStr = sqlStr + " m.ipkumdiv, m.ipkumdate, m.accountdiv, m.regdate, m.reqphone, m.reqhp, m.deliverno, "
		sqlStr = sqlStr + " m.sitename, m.discountrate, m.cancelyn, "
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoption, d.itemno, d.itemoptionname as optname, d.itemcost,"
		sqlStr = sqlStr + " d.beasongdate,d.isupchebeasong, d.buycash, d.currstate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.sitename='diyitem'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1'"             ''2009 변경, 주문접수건도 표시
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by d.detailidx desc"

		rsACADEMYget.pagesize = FPageSize

		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		'rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
    		do until rsACADEMYget.eof
    			set FMasterItemList(i) = new CDesignerJumunList
    			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
    			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
    			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
    			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
    			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
    			FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
    			FMasterItemList(i).Fregdate		= rsACADEMYget("regdate")
    			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
    			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
    			FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
    			FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
    			FMasterItemList(i).Fdeliverno	= rsACADEMYget("deliverno")
    			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
    			FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
    			FMasterItemList(i).FCancelyn	= rsACADEMYget("cancelyn")

    			FMasterItemList(i).FItemID       = rsACADEMYget("itemid")
    			FMasterItemList(i).FItemName     = db2Html(rsACADEMYget("itemname"))
    			FMasterItemList(i).FItemOption   = rsACADEMYget("itemoption")
    			FMasterItemList(i).FItemOptionStr= db2Html(rsACADEMYget("optname"))
    			FMasterItemList(i).FItemNo     = rsACADEMYget("itemno")
    			FMasterItemList(i).Fitemcost     = rsACADEMYget("itemcost")

    			FMasterItemList(i).FUpcheBaesongDate     = rsACADEMYget("beasongdate")
    			FMasterItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
    			FMasterItemList(i).Fbuycash = rsACADEMYget("buycash")
    			FMasterItemList(i).FCurrState		 = rsACADEMYget("currstate")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public Sub SearchOnlyOnJumunList()
		dim sqlStr,i
		dim wheredetail

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


		''#################################################
		''총 갯수. 총금액
		''#################################################
		sqlStr = "select count(T.orderserial) as cnt, sum(T.subtotalprice) as subtotal , avg(T.subtotalprice) as avgtotal "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " select m.orderserial, m.subtotalprice,  count(d.detailidx) as dcnt "
		sqlStr = sqlStr + "  from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + "  ,[db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by m.orderserial, m.subtotalprice"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where T.dcnt=1"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")

		FSubtotal = rsACADEMYget("subtotal")
		FAvgTotal = rsACADEMYget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + "T.* "
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " select  m.idx, m.orderserial, m.jumundiv, "
		sqlStr = sqlStr + " m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, "
		sqlStr = sqlStr + " count(d.detailidx) as dcnt "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m ,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d "
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsACADEMYget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsACADEMYget("cvreg")
			FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
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

		response.write "시스템팀 문의"
		response.end

		''#################################################
		''총 갯수. 총금액
		''#################################################
		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where m.orderserial in ("
		sqlStr = sqlStr + " select top 100 m.orderserial"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " and d.itemid=0"
		sqlStr = sqlStr + " and d.itemoption<>'0101'"
		sqlStr = sqlStr + " and d.itemoption<>'0501'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by m.orderserial desc"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + wheredetail
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")

		FSubtotal = rsACADEMYget("subtotal")
		FAvgTotal = rsACADEMYget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where m.orderserial in ("
		sqlStr = sqlStr + " select top 100 m.orderserial"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsACADEMYget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsACADEMYget("cvreg")
			FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))
			FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsACADEMYget("miletotalprice"))
			FMasterItemList(i).Fjungsanflag		= rsACADEMYget("jungsanflag")
			FMasterItemList(i).Freqzipaddr		= db2Html(rsACADEMYget("reqzipaddr"))
			FMasterItemList(i).Fauthcode		= rsACADEMYget("authcode")

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
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

		response.write "시스템팀 문의"
		response.end

		''#################################################
		''총 갯수. 총금액
		''#################################################
		if Fnotitemlist<>"" then
			sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_academy].[dbo].tbl_academy_order_master m"
			sqlStr = sqlStr + " where orderserial not in ("
			sqlStr = sqlStr + " select distinct m.orderserial from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fnotitemlist + ")"
			sqlStr = sqlStr + " )"
			sqlStr = sqlStr + wheredetail
		elseif Fitemlist<>"" then
			sqlStr = "selec11t count(distinct m.orderserial) as cnt, sum(distinct m.subtotalprice) as subtotal , avg(distinct m.subtotalprice) as avgtotal  from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fitemlist + ")"
		else
			sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  from [db_academy].[dbo].tbl_academy_order_master m"
			sqlStr = sqlStr + " where m.idx<>0"
			sqlStr = sqlStr + wheredetail

		end if
'response.write sqlStr
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")

		FSubtotal = rsACADEMYget("subtotal")
		FAvgTotal = rsACADEMYget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		if Fnotitemlist<>"" then
			sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg from [db_academy].[dbo].tbl_academy_order_master m"
			sqlStr = sqlStr + " where orderserial not in ("
			sqlStr = sqlStr + " select distinct m.orderserial from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
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

			sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + Fitemlist + ")"
			sqlStr = sqlStr + " order by m.idx "
		else
			sqlStr = "select top " + CStr(FPageSize) + " *, convert(varchar,m.regdate,20) as cvreg"
			sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
			sqlStr = sqlStr + " where m.idx<>0"
			sqlStr = sqlStr + wheredetail
			sqlStr = sqlStr + " order by idx "


		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
			set FMasterItemList(i) = new CJumunMasterItem
			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountname	= db2Html(rsACADEMYget("accountname"))
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Faccountno	= rsACADEMYget("accountno")
			FMasterItemList(i).Ftotalvat	= Null2Zoro(rsACADEMYget("totalvat"))
			FMasterItemList(i).Ftotalmileage= Null2Zoro(rsACADEMYget("totalmileage"))
			FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
			FMasterItemList(i).Fregdate		= rsACADEMYget("cvreg")
			FMasterItemList(i).Fbeadaldiv	= rsACADEMYget("beadaldiv")
			FMasterItemList(i).Fbeadaldate	= rsACADEMYget("beadaldate")
			FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Fbuyphone	= rsACADEMYget("buyphone")
			FMasterItemList(i).Fbuyhp		= rsACADEMYget("buyhp")
			FMasterItemList(i).Fbuyemail	= rsACADEMYget("buyemail")
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Freqzipcode	= rsACADEMYget("reqzipcode")
			FMasterItemList(i).Freqaddress	= db2Html(rsACADEMYget("reqaddress"))
			FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
			FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
			''/FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
			FMasterItemList(i).Fdeliverno	= rsACADEMYget("deliverno")
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fpaygatetid	= rsACADEMYget("paygatetid")
			FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
			FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))
			FMasterItemList(i).Fresultmsg		= rsACADEMYget("resultmsg")
			FMasterItemList(i).Frduserid		= rsACADEMYget("rduserid")
			FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsACADEMYget("miletotalprice"))
			FMasterItemList(i).Fjungsanflag		= rsACADEMYget("jungsanflag")
			FMasterItemList(i).Freqzipaddr		= db2Html(rsACADEMYget("reqzipaddr"))
			FMasterItemList(i).Fauthcode		= rsACADEMYget("authcode")

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
	end sub

	public Sub SearchJumunList()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectUserID<>"") then
			wheredetail = wheredetail + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			wheredetail = wheredetail + " and buyname = '" + FRectBuyname + "'"
		end if

		if (FRectReqName<>"") then
			wheredetail = wheredetail + " and reqname = '" + FRectReqName + "'"
		end if


		if (FRectSubTotalPrice<>"") then
			wheredetail = wheredetail + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and cancelyn ='N'"
		end if

        if (FRectIpkumdiv<>"") then
            wheredetail = wheredetail + " and ipkumdiv='" & FRectIpkumdiv & "'"
        end if

		if (FRectIpkumDiv2<>"") then
			wheredetail = wheredetail + " and ipkumdiv>='2'"
		end if

		if (FRectIpkumDiv4<>"") then
			wheredetail = wheredetail + " and ipkumdiv>='4'"
		end if

		if (FRectOnlyIpkumDiv<>"") then
			wheredetail = wheredetail + " and ipkumdiv=" + CStr(FRectOnlyIpkumDiv)
		end if

		if (FRectIpkumName<>"") then
			wheredetail = wheredetail + " and accountname = '" + FRectIpkumName + "'"
		end if

		if (FRectSiteName<>"") then
			wheredetail = wheredetail + " and sitename ='" + FRectSiteName + "'"
		end if

		if (FRectRdSite<>"") then
			wheredetail = wheredetail + " and rdsite ='" + FRectRdSite + "'"
		end if

		if (FRectNoViewPoint<>"") then
			wheredetail = wheredetail + " and accountdiv<>'30'"
		end if

		if (FRectOnlyOutMall<>"") then
			wheredetail = wheredetail + " and accountdiv='50'"
		end if

		if (FRectOnlyPoint<>"") then
			wheredetail = wheredetail + " and accountdiv='30'"
		end if

        if (FRectIsFlower="Y") then
			wheredetail = wheredetail + " and cardribbon is Not NULL "
		end if
		if (FRectIsMinus="Y") then
			wheredetail = wheredetail + " and jumundiv='9' "
		end if
		if (FRectIsForeign<>"") then
            wheredetail = wheredetail + " and ((IsNULL(dlvcountryCode,'KR')<>'KR') and (IsNULL(dlvcountryCode,'KR')<>'ZZ')) "
        end if
		if (FRectIsMilitary<>"") then
            wheredetail = wheredetail + " and (IsNULL(dlvcountryCode,'KR') = 'ZZ') "
        end if

		''#################################################
		''총 갯수. 총금액
		''#################################################
		sqlStr = "select count(orderserial) as cnt, sum(subtotalprice) as subtotal , avg(subtotalprice) as avgtotal"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master "
		sqlStr = sqlStr + " where idx<>0"
		sqlStr = sqlStr + wheredetail
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")

		FSubtotal = rsACADEMYget("subtotal")
		FAvgTotal = rsACADEMYget("avgtotal")

		if IsNull(FSubtotal) then FSubtotal=0
		if IsNull(FAvgTotal) then FAvgTotal=0
		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " *, convert(varchar,regdate,20) as cvreg, convert(varchar,ipkumdate,20) as cvipkumdate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master "
		sqlStr = sqlStr + " where idx<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)
		i=0
		if not rsACADEMYget.Eof then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FMasterItemList(i) = new CJumunMasterItem
				FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
				FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
				FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
				FMasterItemList(i).Faccountname	= db2Html(rsACADEMYget("accountname"))
				FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
				FMasterItemList(i).Faccountno	= rsACADEMYget("accountno")
				'FMasterItemList(i).Ftotalvat	= Null2Zoro(rsACADEMYget("totalvat"))
				FMasterItemList(i).Ftotalmileage= Null2Zoro(rsACADEMYget("totalmileage"))
				FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
				FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
				FMasterItemList(i).Fipkumdate	= rsACADEMYget("cvipkumdate")
				FMasterItemList(i).Fregdate		= rsACADEMYget("cvreg")
				FMasterItemList(i).Fbeadaldiv	= rsACADEMYget("songjangdiv")
				FMasterItemList(i).Fbeadaldate	= rsACADEMYget("beadaldate")
				FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
				FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
				FMasterItemList(i).Fbuyphone	= rsACADEMYget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsACADEMYget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsACADEMYget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
				FMasterItemList(i).Freqzipcode	= rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqaddress	= db2Html(rsACADEMYget("reqaddress"))
				FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
				FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
				FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
				FMasterItemList(i).Fdeliverno	= rsACADEMYget("deliverno")
				FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
				FMasterItemList(i).Fpaygatetid	= rsACADEMYget("paygatetid")
				FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
				FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))
				FMasterItemList(i).Fresultmsg		= rsACADEMYget("resultmsg")
				FMasterItemList(i).Frduserid		= rsACADEMYget("rduserid")
				FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsACADEMYget("miletotalprice"))
				FMasterItemList(i).Fjungsanflag		= rsACADEMYget("jungsanflag")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsACADEMYget("reqzipaddr"))
				FMasterItemList(i).Fauthcode		= rsACADEMYget("authcode")
				FMasterItemList(i).Fcouponpay	    = rsACADEMYget("tencardspend")

				FMasterItemList(i).Fcardribbon	= rsACADEMYget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsACADEMYget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsACADEMYget("fromname"))
				FMasterItemList(i).Freqdate  	= rsACADEMYget("reqdate")
				FMasterItemList(i).Freqtime 	= db2html(rsACADEMYget("reqtime"))

				FMasterItemList(i).Fuserlevel	= rsACADEMYget("userlevel")
                'FMasterItemList(i).FDlvcountryCode = rsACADEMYget("DlvcountryCode")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub


	public Sub SearchMatchJumunList()
		dim sqlStr
		dim wheredetail
		dim i

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


		if (FRectSearchtype02<>"") then
			wheredetail = wheredetail + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and cancelyn ='N'"
		end if

		if (FRectIpkumDiv2<>"") then
			wheredetail = wheredetail + " and ipkumdiv=2"
		else
			wheredetail = wheredetail + " and ipkumdiv>=2"
		end if

		wheredetail = wheredetail + " and accountdiv='7'"



		''#################################################
		''총 갯수. 총금액
		''#################################################
''		sqlStr = "select count(orderserial) as cnt, sum(subtotalprice) as subtotal , avg(subtotalprice) as avgtotal"
''		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master "
''		sqlStr = sqlStr + " where idx<>0"
''		sqlStr = sqlStr + wheredetail
''		rsACADEMYget.Open sqlStr,dbACADEMYget,1
''		FTotalCount = rsACADEMYget("cnt")
''
''		FSubtotal = rsACADEMYget("subtotal")
''		FAvgTotal = rsACADEMYget("avgtotal")
''
''		if IsNull(FSubtotal) then FSubtotal=0
''		if IsNull(FAvgTotal) then FAvgTotal=0
''		rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " *, convert(varchar,regdate,20) as cvreg"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by orderserial desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)
		i=0
		if not rsACADEMYget.Eof then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FMasterItemList(i) = new CJumunMasterItem
				FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
				FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
				FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
				FMasterItemList(i).Faccountname	= db2Html(rsACADEMYget("accountname"))
				FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
				FMasterItemList(i).Faccountno	= rsACADEMYget("accountno")
				FMasterItemList(i).Ftotalvat	= Null2Zoro(rsACADEMYget("totalvat"))
				FMasterItemList(i).Ftotalmileage= Null2Zoro(rsACADEMYget("totalmileage"))
				FMasterItemList(i).Ftotalsum	= rsACADEMYget("totalsum")
				FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
				FMasterItemList(i).Fipkumdate	= rsACADEMYget("ipkumdate")
				FMasterItemList(i).Fregdate		= rsACADEMYget("cvreg")
				FMasterItemList(i).Fbeadaldiv	= rsACADEMYget("beadaldiv")
				FMasterItemList(i).Fbeadaldate	= rsACADEMYget("beadaldate")
				FMasterItemList(i).Fcancelyn	= rsACADEMYget("cancelyn")
				FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
				FMasterItemList(i).Fbuyphone	= rsACADEMYget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsACADEMYget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsACADEMYget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
				FMasterItemList(i).Freqzipcode	= rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqaddress	= db2Html(rsACADEMYget("reqaddress"))
				FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
				FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
				FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
				FMasterItemList(i).Fdeliverno	= rsACADEMYget("deliverno")
				FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
				FMasterItemList(i).Fpaygatetid	= rsACADEMYget("paygatetid")
				FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
				FMasterItemList(i).Fsubtotalprice	= Null2Zoro(rsACADEMYget("subtotalprice"))
				FMasterItemList(i).Fresultmsg		= rsACADEMYget("resultmsg")
				FMasterItemList(i).Frduserid		= rsACADEMYget("rduserid")
				FMasterItemList(i).Fmiletotalprice	= Null2Zoro(rsACADEMYget("miletotalprice"))
				FMasterItemList(i).Fjungsanflag		= rsACADEMYget("jungsanflag")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsACADEMYget("reqzipaddr"))
				FMasterItemList(i).Fauthcode		= rsACADEMYget("authcode")
				FMasterItemList(i).Fcouponpay	    = rsACADEMYget("tencardspend")

				FMasterItemList(i).Fcardribbon	= rsACADEMYget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsACADEMYget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsACADEMYget("fromname"))
				FMasterItemList(i).Freqdate  	= rsACADEMYget("reqdate")
				FMasterItemList(i).Freqtime 	= db2html(rsACADEMYget("reqtime"))

				FMasterItemList(i).Fuserlevel	= rsACADEMYget("userlevel")


				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

        '빈 주문데이타(orderserial 이 없는, 빈 디테일 페이지 표시에 사용)
	public Sub GetBlankOneJumunList()
		dim sqlStr
		dim wheredetail
		dim i

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
		dim sqlStr
		dim i


		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select "
		sqlStr = sqlStr + " d.itemid, d.buycash ,d.itemcost, sum(d.itemno) as sm, d.itemname, d.itemoptionname"
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
			sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_detail d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		end if
		'sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, "
		'sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
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
		sqlStr = sqlStr + " and m.sitename = 'diyitem'"
		sqlStr = sqlStr + " and d.itemid <> 0"

		sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.buycash, d.itemcost, d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " order by d.itemid desc"

		rsACADEMYget.PageSize = FPageSize

		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		''rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsACADEMYget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsACADEMYget("sm")
				FMasterItemList(i).FItemID       = rsACADEMYget("itemid")
				FMasterItemList(i).FItemCost     = rsACADEMYget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsACADEMYget("itemoptionname"))
				FMasterItemList(i).FBuycash		= rsACADEMYget("buycash")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public Sub SearchJumunListByupcheSelllist2()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""


		if (FRectRegStart<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			elseif (FRectDateType="beadal") then
				wheredetail = wheredetail + " and m.beadaldate >='" + CStr(FRectRegStart) + "'"
			else
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif (FRectDateType="beadal") then
				wheredetail = wheredetail + " and m.beadaldate <'" + CStr(FRectRegEnd) + "'"
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





		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " d.itemid, d.itemoption, sum(d.itemno) as sm, "
		sqlStr = sqlStr + " d.itemcost,d.buycash,d.itemname,d.itemoptionname"

		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
			sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_detail d "
			sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		end if
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.sitename='diyitem'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0 and m.ipkumdiv>1"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.itemcost,d.buycash,d.itemname,d.itemoptionname"
		sqlStr = sqlStr + " order by sm desc, d.itemid, d.itemoption"

		rsACADEMYget.PageSize = FPageSize

		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FTotalCount = rsACADEMYget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsACADEMYget("sm")
				FMasterItemList(i).FItemID       = rsACADEMYget("itemid")
				FMasterItemList(i).FItemCost       = rsACADEMYget("itemcost")
				FMasterItemList(i).Fbuycash       = rsACADEMYget("buycash")
				FMasterItemList(i).FItemName     = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsACADEMYget("itemoptionname"))

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
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


		''#################################################
		''총 갯수. 총금액
		''#################################################
		'sqlStr = "select count(d.orderserial) as cnt"
		'sqlStr = sqlStr + " from tbl_academy_order_master m,"
		'sqlStr = sqlStr + " tbl_academy_order_detail d, tbl_diy_item i"
		'sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		'sqlStr = sqlStr + wheredetail
		'sqlStr = sqlStr + " and d.itemid=i.itemid"
		'sqlStr = sqlStr + " and i.makerid='" + FRectDesignerID + "'"

		'rsACADEMYget.Open sqlStr,dbACADEMYget,1
		'FTotalCount = rsACADEMYget("cnt")

		'if IsNull(FSubtotal) then FSubtotal=0
		'if IsNull(FAvgTotal) then FAvgTotal=0
		'rsACADEMYget.Close

		''#################################################
		''데이타.
		''#################################################

		sqlStr = "select"
		sqlStr = sqlStr + " d.orderserial, m.buyname,m.reqname, m.jumundiv, m.userid,"
		sqlStr = sqlStr + " m.ipkumdiv, m.accountdiv, m.regdate, m.reqphone, m.reqhp, m.deliverno, "
		sqlStr = sqlStr + " m.sitename, m.discountrate, m.cancelyn, "
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoption, d.itemoptionname, d.isupchebeasong"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
		sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.sitename='diyitem'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by m.regdate desc"

'response.write sqlStr

		rsACADEMYget.PageSize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsACADEMYget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage
			do until (i >= FResultCount)

			set FMasterItemList(i) = new CDesignerJumunList

			FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
			FMasterItemList(i).Fjumundiv	= rsACADEMYget("jumundiv")
			FMasterItemList(i).Fuserid		= rsACADEMYget("userid")
			FMasterItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FMasterItemList(i).Fipkumdiv	= rsACADEMYget("ipkumdiv")
			FMasterItemList(i).Fregdate		= rsACADEMYget("regdate")
			FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
			FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
			FMasterItemList(i).Fdeliverno	= rsACADEMYget("deliverno")
			FMasterItemList(i).Fdeliverytype	= rsACADEMYget("isupchebeasong")
			FMasterItemList(i).Fsitename	= rsACADEMYget("sitename")
			FMasterItemList(i).Fdiscountrate	= rsACADEMYget("discountrate")
			FMasterItemList(i).FCancelyn	= rsACADEMYget("cancelyn")

			FMasterItemList(i).FItemID       = rsACADEMYget("itemid")
			FMasterItemList(i).FItemName     = db2html(rsACADEMYget("itemname"))
			FMasterItemList(i).FItemOption   = rsACADEMYget("itemoption")
			FMasterItemList(i).FItemOptionStr= db2html(rsACADEMYget("itemoptionname"))

			rsACADEMYget.movenext
			i = i + 1
		loop
	end if
		rsACADEMYget.Close
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
