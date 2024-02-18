<%
'###########################################################
' Description : 제휴몰 다음
' Hieditor : 2015.05.27 김진영 생성
'			 2016.07.21 한용민 수정
'###########################################################

Class CJumunMasterItem
	Public Forderserial
	Public Fjumundiv
	Public Fuserid
	Public Ftotalmileage
	Public Ftotalsum
	Public Fipkumdiv
	Public Fipkumdate
	Public Fregdate
	Public Fbeadaldate
	Public Fcancelyn
	Public Fsitename
	Public Fsubtotalprice

	Public Fmiletotalprice
	Public Fjungsanflag

	Public FDtlItemName
	Public FDtlItemNo
	Public FDtlItemOption
	Public FDtlItemOptionName

    Public FsumpaymentEtc       '''''2011-04 추가

    public Fcanceldate

    public FTenCardSpend
	public FAllatDiscountPrice
    public FitemcostSum
    public FreducedpriceSum
    public FtargetNoVatSum
    public FdlvcostSum
    public FdlvcostCpnSum
    public Frdsite

    ''Public FdlvPaySum           ''''2013/09/24

    public function getRdSiteName()
        getRdSiteName=Frdsite

        if isNULL(Frdsite) then  Exit function

'        if (isMobileOrder) then
'            getRdSiteName = replace(Frdsite,"mobile_","")
'        end if
    end function

    public function isMobileOrder()
        isMobileOrder = false
        if isNULL(Frdsite) then Exit function

        if Left(Frdsite,7)="mobile_" then
            isMobileOrder=true
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

    public function getDlvPaySum()
        getDlvPaySum = 0
        getDlvPaySum = FdlvcostSum-FdlvcostCpnSum
    end function

         
'    public function getPriceBonusCouponSum
'        getPriceBonusCouponSum = 0
'        if FTenCardSpend=0 then Exit function
'        if FdlvcostCpnSum<>0 then Exit function
'
'        getPriceBonusCouponSum=FTenCardSpend-FdlvcostCpnSum-(FitemcostSum-FreducedpriceSum)
'    end function

    public function getEnuiSum()
        getEnuiSum = 0
        getEnuiSum = CLNG(FTenCardSpend)+CLNG(FAllatDiscountPrice) ''-CLNG(FdlvcostCpnSum)
    end function

    public function getJungsanTargetNoVatSum
        getJungsanTargetNoVatSum = FtargetNoVatSum ''getJungsanTargetSum*10/11
    end function

'    public function getJungsanTargetSum()
'        getJungsanTargetSum =0
'        if (FreducedpriceSum=0) then Exit function
'        getJungsanTargetSum =FreducedpriceVatSum/FreducedpriceSum*(FreducedpriceSum-getPriceBonusCouponSum)
'
'    end function

    public function isJungsanFixed()
        isJungsanFixed = false
        if isNULL(Fipkumdate) then Exit function
        if (Fipkumdiv<8) then Exit function
        if isNULL(Fbeadaldate) then Exit function

        isJungsanFixed=true
    end function

    public Function getJungsanFixdate() ''정산일
       if (Not isJungsanFixed) then Exit function

       getJungsanFixdate = Left(Fbeadaldate,10)
    end function

    Public Function TotalMajorPaymentPrice()
        TotalMajorPaymentPrice = CLNG(FsubtotalPrice) - CLNG(FsumPaymentEtc)
    End Function

    ''해외배송인지여부
	Public Function IsForeignDeliver()
		IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR")  and (FDlvcountryCode<>"ZZ")
	End Function

	'/사용금지		'/공통펑션에 공용함수 쓸것.		'/2016.07.21 한용민
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

	'/사용금지		'/공통펑션에 공용함수 쓸것.		'/2016.07.21 한용민
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

    public function getCanceldate()
        if Fcancelyn="N"  then Exit function
        if isNULL(Fcanceldate ) then Exit function

        getCanceldate = Left(Fcanceldate,10)
    end function

	public function GetRegDate()
		GetRegDate = Left(FRegDate,10)
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
		elseif Faccountdiv="14" then
			JumunMethodName="편의점결제"
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

Class CJumunMaster
    public FOneItem
	Public FMasterItemList()
	Public FMasterItemList2()
	Public FJumunDetail
	Public FTotalCount
	Public FTotalBuyCash

	Public FSubtotal
	Public FSumTotal

	Public FResultCount

	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectSearchtype
	Public FRectSearchtype01
	Public FRectSearchtype02
	Public FRectOrderSerial
	Public FRectUserID
	Public FRectckdate
	Public FRectBuyname
	Public FRectReqName
	Public FRectIpkumName
	Public FRectSubTotalPrice

	Public FRectRegStart
	Public FRectRegEnd
	Public FRectDelNoSearch
	Public FRectIpkumDiv2
	Public FRectIpkumDiv4
	Public FRectMeCode
	Public FRectSiteName
	Public FRectRdSite
	Public FRectOnlyIpkumDiv
    Public FRectckpointsearch

	Public FRectNotThisSite
	Public FRectNoViewPoint

	Public FRectDesignerID
	Public FRectItemid
	Public FRectItemName
	Public FRectOnlyOutMall
	Public FRectDateType

	Public FRectOrderBy
	Public FRectDispY
	Public FRectSellY
	Public FRectOnlyPoint

	Public FRectIpkumOrJumun
	Public FRectBeasongNotFinish
	Public Fnotitemlist
	Public Fitemlist
	Public FRectIpkumDiv4before
	Public FRectOldJumun
	Public FRectcancelyn
    Public FRectIpkumDiv
    Public FRectMinusOrderInclude
	Public FRectreqdate
    Public FRectIsUpcheBeasong
	Public FRectreqHp
    Public FRectIsFlower
    Public FRectIsMinus
    Public FRectIsForeign
    Public FRectIsMilitary

    Public FRectJumunDiv
    Public FRectBuyHp
    Public FRectBuyEmail

    Public FRectCDL
    Public FRectCDM
    Public FRectCDS

    Public FRectBrandPurchaseType
    Public FIsMDPick
    Public FIsRdSite

    Public FRectQryDLVsum
    Public FRectMType
    
	Private Sub Class_Initialize()
		Redim FMasterItemList(0)
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
		GetImageFolerName = GetImageSubFolderByItemid(FJumunDetail.FJumunDetailList(i).FItemID)
	end function

    public Sub GetOneDaumEpJumunMaster
        dim sqlStr
        sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1"
		sqlStr = sqlStr & " m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0) as miletotalprice,m.tencardspend,m.allatdiscountprice"
        sqlStr = sqlStr & " ,m.beadaldate,m.ipkumdiv,m.jumundiv, m.ipkumdate, m.regdate, m.cancelyn, m.canceldate, m.jungsanflag, m.rdsite"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 then d.itemcost*d.itemno else 0 END) as itemcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N'  then d.reducedprice*d.itemno else 0 END) as reducedpriceSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N'  then round((CASE WHEN d.vatinclude='Y' THEN d.reducedprice*d.itemno/1.1 ELSE d.reducedprice*d.itemno END),0) else 0 END) as targetNoVatSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 then d.itemcost*d.itemno else 0 END) as dlvcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 and d.bonuscouponidx is Not NULL then (d.itemcost-d.reducedprice)*d.itemno else 0 END) as dlvcostCpnSum"
        sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr & " 	JOIN db_item.dbo.tbl_Outmall_RdsiteGubun as G "
        sqlStr = sqlStr & " 	on m.rdsite = G.rdsite "
        sqlStr = sqlStr & " 	and G.gubun = 'daumshop'"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on D.orderserial=m.orderserial"
        sqlStr = sqlStr & " 	and D.cancelyn<>'Y'" ''--부분취소 재낌.
        sqlStr = sqlStr & " where m.ipkumdiv>1" ''주문접수 이상
        sqlStr = sqlStr & " and m.jumundiv<>6"
        sqlStr = sqlStr & " and NOT (m.jumundiv='9' and m.cancelyn='D')"
        sqlStr = sqlStr & " and m.orderserial='"&FRectOrderserial&"'"
        sqlStr = sqlStr & " group by m.idx,m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0),m.tencardspend,m.allatdiscountprice,m.beadaldate,m.ipkumdiv,m.jumundiv, m.ipkumdate, m.regdate, m.cancelyn, m.canceldate, m.jungsanflag, m.rdsite"
        sqlStr = sqlStr & " order by m.orderserial desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			set FOneItem = new CJumunMasterItem
				FOneItem.Forderserial		    = rsget("orderserial")
				FOneItem.Fjumundiv		    = rsget("jumundiv")
				FOneItem.Ftotalsum		    = rsget("totalsum")
				FOneItem.Fipkumdiv		    = rsget("ipkumdiv")
				FOneItem.Fipkumdate		    = rsget("ipkumdate")
				FOneItem.Fregdate			    = rsget("regdate")
				FOneItem.Fbeadaldate		    = rsget("beadaldate")
				FOneItem.Fcancelyn		    = rsget("cancelyn")
				FOneItem.Fcanceldate          = rsget("canceldate")
				FOneItem.Fsubtotalprice	    = Null2Zoro(rsget("subtotalprice"))
				FOneItem.Fmiletotalprice	    = Null2Zoro(rsget("miletotalprice"))
				FOneItem.Fjungsanflag		    = rsget("jungsanflag")
	            FOneItem.FTenCardSpend 	    = Null2Zoro(rsget("TenCardSpend"))
	            FOneItem.FAllatDiscountPrice 	= Null2Zoro(rsget("AllatDiscountPrice"))
                FOneItem.FitemcostSum         = Null2Zoro(rsget("itemcostSum"))
                FOneItem.FreducedpriceSum     = Null2Zoro(rsget("reducedpriceSum"))
                FOneItem.FtargetNoVatSum  = Null2Zoro(rsget("targetNoVatSum"))
                FOneItem.FdlvcostSum          = Null2Zoro(rsget("dlvcostSum"))
                FOneItem.FdlvcostCpnSum       = Null2Zoro(rsget("dlvcostCpnSum"))
                FOneItem.Frdsite    = rsget("rdsite")
		End If
		rsget.Close
    end Sub

	Public Sub daumEpJumunList()
		dim sqlStr
		dim wheredetail
		dim i
		wheredetail = ""

		if (FRectOrderSerial<>"") then
			wheredetail = wheredetail + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
		    if (FRectMType="ip") then
		        wheredetail = wheredetail + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
		    elseif (FRectMType="fx") then
		        wheredetail = wheredetail + " and m.beadaldate >='" + CStr(FRectRegStart) + "'"
		    else
			    wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		    end if
		end if

		if (FRectRegEnd<>"") then
		    if (FRectMType="ip") then
		        wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
		    elseif (FRectMType="fx") then   
		        wheredetail = wheredetail + " and m.beadaldate <'" + CStr(FRectRegEnd) + "'"
		    else 
			    wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectMeCode <> "") then
			If FRectMeCode = "WEB_ALL" Then
				wheredetail = wheredetail + " and Left(G.rdsite, 8) = 'daumshop' "
			ElseIf FRectMeCode = "MOBILE_ALL" Then
				wheredetail = wheredetail + " and Left(G.rdsite, 6) = 'mobile' "
			Else
				wheredetail = wheredetail + " and G.rdsite in ('"&FRectMeCode&"')"
			End If
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(orderserial) as cnt, sum(subtotalprice) as subtotalprice, sum(totalsum) as totalsum "
		sqlStr = sqlStr & " ,sum(miletotalprice) as miletotalprice, sum(tencardspend) as tencardspend, sum(allatdiscountprice) as allatdiscountprice"
		sqlStr = sqlStr & " ,sum(itemcostSum) as itemcostSum"
		sqlStr = sqlStr & " ,sum(reducedpriceSum) as reducedpriceSum"
		sqlStr = sqlStr & " ,sum(targetNoVatSum) as targetNoVatSum"
		sqlStr = sqlStr & " ,sum(dlvcostSum) as dlvcostSum"
		sqlStr = sqlStr & " ,sum(dlvcostCpnSum) as dlvcostCpnSum"
		sqlStr = sqlStr & " From ("
		sqlStr = sqlStr & " select m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0) as miletotalprice,m.tencardspend,m.allatdiscountprice"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 then d.itemcost*d.itemno else 0 END) as itemcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N' then d.reducedprice*d.itemno else 0 END) as reducedpriceSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N' then round((CASE WHEN d.vatinclude='Y' THEN d.reducedprice*d.itemno/1.1 ELSE d.reducedprice*d.itemno END),0) else 0 END) as targetNoVatSum"  ''and 
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 then d.itemcost*d.itemno else 0 END) as dlvcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 and d.bonuscouponidx is Not NULL then (d.itemcost-d.reducedprice)*d.itemno else 0 END) as dlvcostCpnSum"
        sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr & " 	JOIN db_item.dbo.tbl_Outmall_RdsiteGubun as G "
        sqlStr = sqlStr & " 	on m.rdsite = G.rdsite "
        sqlStr = sqlStr & " 	and G.gubun in ('daumshop')"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on D.orderserial=m.orderserial"
        sqlStr = sqlStr & " 	and D.cancelyn<>'Y'" ''--부분취소 재낌.
        sqlStr = sqlStr & " where m.ipkumdiv>1" ''주문접수 이상
        sqlStr = sqlStr & " and m.jumundiv<>6"
        sqlStr = sqlStr & " and NOT (m.jumundiv='9' and m.cancelyn='D')"
        sqlStr = sqlStr & wheredetail
        sqlStr = sqlStr & " group by m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0),m.tencardspend,m.allatdiscountprice"
        sqlStr = sqlStr & " ) T"
'response.write sqlStr &"<BR>"
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			FTotalCount	= rsget("cnt")

			set FOneItem = new CJumunMasterItem
			FOneItem.Ftotalsum		    = rsget("totalsum")
			'FOneItem.Fcancelyn		    = rsget("cancelyn")
			'FOneItem.Fcanceldate          = rsget("canceldate")
			FOneItem.Fsubtotalprice	    = Null2Zoro(rsget("subtotalprice"))
			FOneItem.Fmiletotalprice	= Null2Zoro(rsget("miletotalprice"))
			'FOneItem.Fjungsanflag		= rsget("jungsanflag")
            FOneItem.FTenCardSpend 	    = Null2Zoro(rsget("TenCardSpend"))
            FOneItem.FAllatDiscountPrice  = Null2Zoro(rsget("AllatDiscountPrice"))
            FOneItem.FitemcostSum         = Null2Zoro(rsget("itemcostSum"))
            FOneItem.FreducedpriceSum     = Null2Zoro(rsget("reducedpriceSum"))
            FOneItem.FtargetNoVatSum  = Null2Zoro(rsget("targetNoVatSum"))
            FOneItem.FdlvcostSum          = Null2Zoro(rsget("dlvcostSum"))
            FOneItem.FdlvcostCpnSum       = Null2Zoro(rsget("dlvcostCpnSum"))
	    end if
		rsget.Close

        if FTotalCount<1 then Exit Sub

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0) as miletotalprice,m.tencardspend,m.allatdiscountprice"
        sqlStr = sqlStr & " ,m.beadaldate,m.ipkumdiv,m.jumundiv, m.ipkumdate, m.regdate, m.cancelyn, m.canceldate, m.jungsanflag, m.rdsite"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 then d.itemcost*d.itemno else 0 END) as itemcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N' then d.reducedprice*d.itemno else 0 END) as reducedpriceSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid<>0 and d.cancelyn<>'Y' and m.cancelyn='N' then round((CASE WHEN d.vatinclude='Y' THEN d.reducedprice*d.itemno/1.1 ELSE d.reducedprice*d.itemno END),0) else 0 END) as targetNoVatSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 then d.itemcost*d.itemno else 0 END) as dlvcostSum"
        sqlStr = sqlStr & " ,sum(CASE WHEN d.itemid=0 and d.bonuscouponidx is Not NULL then (d.itemcost-d.reducedprice)*d.itemno else 0 END) as dlvcostCpnSum"
        sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr & " 	JOIN db_item.dbo.tbl_Outmall_RdsiteGubun as G "
        sqlStr = sqlStr & " 	on m.rdsite = G.rdsite "
        sqlStr = sqlStr & " 	and G.gubun in ('daumshop')"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on D.orderserial=m.orderserial"
        sqlStr = sqlStr & " 	and D.cancelyn<>'Y'" ''--부분취소 재낌.
        sqlStr = sqlStr & " where m.ipkumdiv>1" ''주문접수 이상
        sqlStr = sqlStr & " and m.jumundiv<>6"
        sqlStr = sqlStr & " and NOT (m.jumundiv='9' and m.cancelyn='D')"
        sqlStr = sqlStr & wheredetail
        sqlStr = sqlStr & " group by m.orderserial,m.totalsum,m.subtotalprice,isNULL(m.miletotalprice,0),m.tencardspend,m.allatdiscountprice,m.beadaldate,m.ipkumdiv,m.jumundiv, m.ipkumdate, m.regdate, m.cancelyn, m.canceldate, m.jungsanflag, m.rdsite"
        sqlStr = sqlStr & " order by m.orderserial desc"

'response.write sqlStr &"<BR>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		If (FTotalCount \ FPageSize) <> (FTotalCount / FPageSize) Then
			FtotalPage = FtotalPage + 1
		End If
		FResultCount = rsget.RecordCount - (FPageSize * (FCurrPage - 1))
		Redim preserve FMasterItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FMasterItemList(i) = new CJumunMasterItem
					FMasterItemList(i).Forderserial		    = rsget("orderserial")
					FMasterItemList(i).Fjumundiv		    = rsget("jumundiv")
					FMasterItemList(i).Ftotalsum		    = rsget("totalsum")
					FMasterItemList(i).Fipkumdiv		    = rsget("ipkumdiv")
					FMasterItemList(i).Fipkumdate		    = rsget("ipkumdate")
					FMasterItemList(i).Fregdate			    = rsget("regdate")
					FMasterItemList(i).Fbeadaldate		    = rsget("beadaldate")
					FMasterItemList(i).Fcancelyn		    = rsget("cancelyn")
					FMasterItemList(i).Fcanceldate          = rsget("canceldate")
					FMasterItemList(i).Fsubtotalprice	    = Null2Zoro(rsget("subtotalprice"))
					FMasterItemList(i).Fmiletotalprice	    = Null2Zoro(rsget("miletotalprice"))
					FMasterItemList(i).Fjungsanflag		    = rsget("jungsanflag")
					FMasterItemList(i).Frdsite              = rsget("rdsite")
		            FMasterItemList(i).FTenCardSpend 	    = Null2Zoro(rsget("TenCardSpend"))
		            FMasterItemList(i).FAllatDiscountPrice 	= Null2Zoro(rsget("AllatDiscountPrice"))
                    FMasterItemList(i).FitemcostSum         = Null2Zoro(rsget("itemcostSum"))
                    FMasterItemList(i).FreducedpriceSum     = Null2Zoro(rsget("reducedpriceSum"))
                    FMasterItemList(i).FtargetNoVatSum  = Null2Zoro(rsget("targetNoVatSum"))
                    FMasterItemList(i).FdlvcostSum          = Null2Zoro(rsget("dlvcostSum"))
                    FMasterItemList(i).FdlvcostCpnSum       = Null2Zoro(rsget("dlvcostCpnSum"))


				rsget.movenext
				i = i + 1
			Loop
		End If
		rsget.Close
	End Sub

	public sub SearchJumunDetail(byval orderserial)
		dim sqlStr
		dim i
		sqlStr = "select top 300 d.idx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.reducedprice,d.mileage,d.cancelyn,"
		sqlStr = sqlStr + " d.itemname, d.makerid, i.listimage as imglist, i.smallimage as imgsmall, d.itemoptionname as codeview"
		sqlStr = sqlStr + " , d.currstate, d.songjangdiv, d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem, d.requiredetail  "
		sqlStr = sqlStr + " , d.vatinclude "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid  "
		sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'"
		sqlStr = sqlStr + " order by d.itemid,d.itemoption"

		rsget.Open sqlStr,dbget,1
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
			FJumunDetail.FJumunDetailList(i).Freducedprice     = rsget("reducedprice")
			FJumunDetail.FJumunDetailList(i).Fmileage     = rsget("mileage")
			FJumunDetail.FJumunDetailList(i).Fcancelyn    = rsget("cancelyn")

			FJumunDetail.FJumunDetailList(i).FItemName    = db2html(rsget("itemname"))
			FJumunDetail.FJumunDetailList(i).FImageList    = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName(i) + "/" + rsget("imglist")
			FJumunDetail.FJumunDetailList(i).FImageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")

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
			FJumunDetail.FJumunDetailList(i).Fvatinclude = rsget("vatinclude")

			rsget.movenext
			i=i+1
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
End Class

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

class CJumunDetailItem
	public Forderserial
	public Fdetailidx

	public FMakerid
	public Fitemid
	public Fitemoption
	public Fitemno
	public Fitemcost
	public Freducedprice
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
	Public Fvatinclude
    public FOmwDiv
    public FoDlvType

    public function getDtlStateName()
        if (Fcancelyn="Y") then
            getDtlStateName="취소"
        else
            if (Fcurrstate=7) and Not isNULL(Fbeasongdate) then
                getDtlStateName = "출고완료"
            elseif (Fcurrstate=3) then
                getDtlStateName = "상품준비"
            end if
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
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>