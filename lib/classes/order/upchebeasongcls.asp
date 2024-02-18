<%
'###########################################################
' Description : 미출고리스트
' Hieditor : 이상구 생성
'###########################################################

Class CUpCheSMSItem
	public FMakerid
	public FCompanyName
	public Fmitongbocnt
	public FMiBalJuCount
	public FMiBeasongCount
	public FLastSendMsgDay
	public FDeliverHp
	public FUserDiv
	public FSocNameKor

	public FNDayMiBaljuCnt
	public FNDayMiBeasongCnt

	public FP_NDayMiBaljuCnt
	public FP_NDayMiBeasongCnt
    public Fcatecode
    public Fcatename

	public function GetMallName()
		if FUserDiv="02" then
			GetMallName = "디자인"
		elseif FUserDiv="03" then
			GetMallName = "플라워"
		elseif FUserDiv="04" then
			GetMallName = "패션"
		elseif FUserDiv="05" then
			GetMallName = "쥬얼리"
		elseif FUserDiv="06" then
			GetMallName = "뷰티"
		elseif FUserDiv="07" then
			GetMallName = "애견"
		elseif FUserDiv="08" then
			GetMallName = "보드게임"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CUpchebeasongDetail
	public FOrderserial
	public FBuyname
	public FReqName
	public FItemID
	public FItemname
	public FItemno
	public FItemoption
	public FItemoptionname
	public FCurrstate
	public FSongjangno
	public FSongjangdiv
	public FIdx
	public FCancelyn
	public FMakerID
	public FOrderDate
	public FIpkumdate
	public FIpkumdiv
	public FDeliverytype
	public FMasterCancel
	public Fdeliverno
	public Fdetailidx

	public Fsitename

	public FItemcnt
	public FJumunDiv

	public FBuyCash
	public FSellcash

	public FUpcheBeasongDate
	public Fmasteridx

	public FRegdate
	public Fbaljudate
	public Fupcheconfirmdate

    public FMisendReason
    public FMisendState
    public FMisendipgodate
    public Fmisendregdate

	public FitemcostCouponNotApplied
	public FOrgitemCost

	public FcsMemoCnt
    public Fomwdiv

	Public Fmisendreguserid

	Public Fmisendmodiuserid
	Public Fmisendmodidate
	Public FsendCount
	Public FlastSendUserid
	Public FlastSendDate
	public FDetailCancelYn

	public Fdlvfinishdt
	public Fjungsanfixdate
	public FDday
    public FDdayByIpkumdate
    public Fvacation

    public function getMisendStateText()
        select Case FMisendState
            CASE 0 : getMisendStateText="미처리"
            CASE 4 : getMisendStateText="고객안내"
            CASE 6 : getMisendStateText="CS처리완료"
            CASE ELSE : getMisendStateText = FMisendState
        end Select
    end function

    public function getMisendText()
    	select Case FMisendReason
			'// 미배송사유
            CASE "00" : getMisendText = "입력대기"

			CASE "03" : getMisendText = "출고지연"
			CASE "02" : getMisendText = "주문제작"
			CASE "08" : getMisendText = "수입"
			CASE "09" : getMisendText = "가구배송"
			CASE "04" : getMisendText = "예약배송"
			CASE "10" : getMisendText = "업체휴가"
			CASE "07" : getMisendText = "고객지정배송"

			CASE "05" : getMisendText = "품절출고불가"
			CASE "66" : getMisendText = "가격오류"

            '' CASE "01" : getMisendText = "재고부족"
            '' CASE "52" : getMisendText = "주문제작"
            '' CASE "53" : getMisendText = "출고지연"
            '' CASE "55" : getMisendText = "품절출고불가"
            CASE ELSE : getMisendText = FMisendReason
        end Select
    end function

    public function getNewBeasongDPlusDateStr()
        getNewBeasongDPlusDateStr = "D+" & FDday
    end function

    public function getBeasongDPlusDateStrByIpkumdate()
        getBeasongDPlusDateStrByIpkumdate = "D+" & FDdayByIpkumdate
    end function

    public function getBeasongDPlusDateStr()
        getBeasongDPlusDateStr = ""

        if IsNULL(Fbaljudate) then
            exit function
        end if

        if IsNULL(FUpcheBeasongDate) then
            getBeasongDPlusDateStr = "D+" & DateDiff("d",Fbaljudate,now())
            exit function
        end if

        getBeasongDPlusDateStr = "D+" & DateDiff("d",Fbaljudate,FUpcheBeasongDate)
    end function

    public function getBeasongDPlusDate()
        getBeasongDPlusDate = ""

        if IsNULL(Fbaljudate) then
            exit function
        end if

        if IsNULL(FUpcheBeasongDate) then
            getBeasongDPlusDate = DateDiff("d",Fbaljudate,now())
            exit function
        end if

        getBeasongDPlusDate = DateDiff("d",Fbaljudate,FUpcheBeasongDate)
    end function

	public function IsAvailAndIpkumOK()
		IsAvailAndIpkumOK = (CInt(Fipkumdiv)>3) and IsAvailJumun
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end class


class CBaljuMaster
	public FMasterItemList()
	public FDetailItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
	public FCurrPage
    public FRectDesignerID
    public FRectItemid
    public FRectItemOption
    public FRectIpkumdiv

	public FRectDateType
	public FRectDeliverType
    public FRect

    public FRectCDL
	public FRectDispCDL
	public FRectDispCDM
	public FRectDispCDS
	Public FRectDispCate
    public FRectDetailState
    public FRectMisendReason
    public FRectMisendState
    public FRectdplusOver
    public FRectdplusLower
    public FRectSiteName
    public FRectSortBy
    public FRectExInMayChulgoDay
    public FRectExInNeedChulgoDay
	public FRectExStockOut
	Public FRectExToday
	Public FRectUpcheNoCheck
	Public FRectCheckMinus
    public FRectIncIpkumdiv4

    public FRectCurrState	' 상태
    public FRectBrandPurchaseType  ''구매형태

    public FOrderCnt
	public FSumItemNo
    public FSumItemCost
    public FSumBuyCash
	public FRectSellChannelDiv
	public frectdetailcancelyn
	public FRectInc3pl
	public FRectchknotcash
	public FRectIsPlusSaleItem
	public FRectIsSendGift

	Private Sub Class_Initialize()

		redim  FMasterItemList(0)
		redim  FDetailItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

    public function GetBaljuPassedDate()
        GetBaljuPassedDate = 0

        if IsNULL(Fbaljudate) then Exit function

        if (Fbaljudate="") then Exit function

        GetBaljuPassedDate = DateDiff("d",(left(Fbaljudate,10)) , (left(now(),10)) )
    end function

	public function IsAvailAndIpkumOK()
		IsAvailAndIpkumOK = (CInt(Fipkumdiv)>3) and IsAvailJumun
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	public Sub DesignerJumunUpcheBeasongFinFind()
		dim sqlStr,i

		sqlStr = " select top 1000 m.idx as midx, m.orderserial, m.buyname, m.reqname, m.regdate, m.ipkumdiv,"
		sqlStr = sqlStr + " m.cancelyn as mastercancel ,m.deliverno, d.itemid, "
		sqlStr = sqlStr + " d.itemname, d.itemno, d.itemoption, d.itemoptionname,"
		sqlStr = sqlStr + " d.currstate, d.songjangno, d.songjangdiv, d.makerid, d.idx,"
		sqlStr = sqlStr + " d.cancelyn "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d, "

		sqlStr = sqlStr + " (select distinct top 3000 m.orderserial,"
		sqlStr = sqlStr + " sum("
		sqlStr = sqlStr + " case d.isupchebeasong"
		sqlStr = sqlStr + " 	when 'Y' then 1"
		sqlStr = sqlStr + " 	else 0"
		sqlStr = sqlStr + " end) as ucnt,"
		sqlStr = sqlStr + " sum("
		sqlStr = sqlStr + " case d.currstate"
		sqlStr = sqlStr + " 	when 7 then 1"
		sqlStr = sqlStr + " 	else 0"
		sqlStr = sqlStr + " end) as scnt,"
		sqlStr = sqlStr + " count(d.idx) as tcnt"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate>='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate<'" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and m.ipkumdiv>=5"
		sqlStr = sqlStr + " and m.ipkumdiv<8"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by m.orderserial"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " where T.ucnt>0 "
		sqlStr = sqlStr + " and T.ucnt=T.tcnt"
		sqlStr = sqlStr + " and T.ucnt=T.scnt"
		sqlStr = sqlStr + " and m.orderserial=T.orderserial"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
 		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and m.ipkumdiv > 4"
		sqlStr = sqlStr + " and d.itemid <>0"
		sqlStr = sqlStr + " order by m.idx desc , d.itemid asc"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1


		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FDetailItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FDetailItemList(i) = new CUpchebeasongDetail

				FDetailItemList(i).FOrderserial    = rsget("orderserial")
				FDetailItemList(i).FBuyname        = db2html(rsget("buyname"))
				FDetailItemList(i).FReqName        = db2html(rsget("reqname"))
				FDetailItemList(i).FItemID         = rsget("itemid")
				FDetailItemList(i).FItemname       = db2html(rsget("itemname"))
				FDetailItemList(i).FItemno         = rsget("itemno")
				FDetailItemList(i).FItemoption     = rsget("itemoption")
				FDetailItemList(i).FItemoptionname = db2html(rsget("itemoptionname"))
				FDetailItemList(i).FCurrstate      = rsget("currstate")
				FDetailItemList(i).FSongjangno     = rsget("songjangno")
				FDetailItemList(i).FSongjangdiv    = rsget("songjangdiv")
				FDetailItemList(i).FIdx            = rsget("idx")
				FDetailItemList(i).FCancelyn       = rsget("cancelyn")
				FDetailItemList(i).FMakerID       = rsget("makerid")
				FDetailItemList(i).FOrderDate		= rsget("regdate")
				FDetailItemList(i).FIpkumdiv		= rsget("ipkumdiv")
				FDetailItemList(i).FMasterCancel    = rsget("mastercancel")
				FDetailItemList(i).Fdeliverno		= rsget("deliverno")

				FDetailItemList(i).Fmasteridx		= rsget("midx")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub DesignerJumunUpcheBeasong()
		dim sqlStr,i

		sqlStr = " select top 2000 m.orderserial, m.buyname, m.reqname, m.regdate, m.ipkumdiv, m.cancelyn as mastercancel"
		sqlStr = sqlStr + " ,m.deliverno, d.itemid, "
		sqlStr = sqlStr + " d.itemname, d.itemno, d.itemoption, d.itemoptionname,"
		sqlStr = sqlStr + " d.currstate, d.songjangno, d.songjangdiv, d.makerid, d.idx, d.cancelyn, i.deliverytype "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " (select distinct top 500 m.orderserial "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate>='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate<'" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
		if FRect="all" then

		elseif FRect="mi" then
			sqlStr = sqlStr + " and not (d.currstate=7 and m.ipkumdiv = 7)"
		else
			sqlStr = sqlStr + " and d.currstate=7 and m.ipkumdiv = 5"
		end if

		sqlStr = sqlStr + " order by m.orderserial desc"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " where m.orderserial=T.orderserial"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
 		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and m.ipkumdiv > 4"
		sqlStr = sqlStr + " and d.itemid <>0"
		sqlStr = sqlStr + " order by m.idx desc , d.itemid asc"
		rsget.Open sqlStr,dbget,1


		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FDetailItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FDetailItemList(i) = new CUpchebeasongDetail

				FDetailItemList(i).FOrderserial    = rsget("orderserial")
				FDetailItemList(i).FBuyname        = db2html(rsget("buyname"))
				FDetailItemList(i).FReqName        = db2html(rsget("reqname"))
				FDetailItemList(i).FItemID         = rsget("itemid")
				FDetailItemList(i).FItemname       = db2html(rsget("itemname"))
				FDetailItemList(i).FItemno         = rsget("itemno")
				FDetailItemList(i).FItemoption     = rsget("itemoption")
				FDetailItemList(i).FItemoptionname = db2html(rsget("itemoptionname"))
				FDetailItemList(i).FCurrstate      = rsget("currstate")
				FDetailItemList(i).FSongjangno     = rsget("songjangno")
				FDetailItemList(i).FSongjangdiv    = rsget("songjangdiv")
				FDetailItemList(i).FIdx            = rsget("idx")
				FDetailItemList(i).FCancelyn       = rsget("cancelyn")
				FDetailItemList(i).FMakerID       = rsget("makerid")
				FDetailItemList(i).FOrderDate		= rsget("regdate")
				FDetailItemList(i).FIpkumdiv		= rsget("ipkumdiv")
				FDetailItemList(i).FDeliverytype    = rsget("deliverytype")
				FDetailItemList(i).FMasterCancel    = rsget("mastercancel")
				FDetailItemList(i).Fdeliverno		= rsget("deliverno")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub DesignerDateSellList()
		dim sqlStr,wheredetail
		dim i
		wheredetail = ""

		if (FRectDesignerID <>"") then
			wheredetail = wheredetail + " and d.makerid='" & FRectDesignerID & "'"
		end if

		if (FRectItemid <>"") then
			wheredetail = wheredetail +  " and d.itemid='" & FRectItemid & "'"
		end if

		if (FRectSiteName <>"") then
			wheredetail = wheredetail +  " and m.sitename='" & FRectSiteName & "'"
		end if

		if (FRectDateType="ipkumil") then
			'// 결제일기준
			wheredetail = wheredetail + " and m.ipkumdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.ipkumdate < '" & FRectRegEnd & "'"
		elseif (FRectDateType="chulgoil") then
			'// 출고일기준(배송비제외)
			wheredetail = wheredetail + " and d.beasongdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and d.beasongdate < '" & FRectRegEnd & "'"
			wheredetail = wheredetail + " and d.itemid <> 0 "
		elseif (FRectDateType="baesongil") then
			'// 배송일기준(배송비제외)
			wheredetail = wheredetail + " and d.dlvfinishdt >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and d.dlvfinishdt < '" & FRectRegEnd & "'"
			wheredetail = wheredetail + " and d.itemid <> 0 "
		elseif (FRectDateType="jungsanil") then
			'// 정산일기준(배송비제외)
			wheredetail = wheredetail + " and d.jungsanfixdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and d.jungsanfixdate < '" & FRectRegEnd & "'"
			wheredetail = wheredetail + " and d.itemid <> 0 "
		else
			'// 주문일기준
			wheredetail = wheredetail + " and m.regdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.regdate < '" & FRectRegEnd & "'"
		end if

		if (FRectDeliverType="upche") then
			wheredetail = wheredetail + " and d.isupchebeasong='Y'"

		elseif (FRectDeliverType="ten") then
			wheredetail = wheredetail + " and d.isupchebeasong<>'Y'"
		else

		end if

		if (FRectDispCate<>"") then
			wheredetail = wheredetail + " and exists(Select 1 from db_item.dbo.tbl_display_cate_item as c WITH(NOLOCK) where c.isDefault='y' " &_
									" and c.itemid=d.itemid " &_
									" and c.catecode like '" & FRectDispCate & "%')"
		end if

		if (FRectCheckMinus="Y") then
			wheredetail = wheredetail + " and (d.itemcost-d.buycash)<=0 "
		end if

        if (FRectSellChannelDiv<>"") then
            wheredetail = wheredetail + " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
        end if

		''2014/01/15추가
		if (FRectInc3pl<>"") then
			if (FRectInc3pl="A") then

			else
				wheredetail = wheredetail & " and isNULL(p2.tplcompanyid,'')<>''"
			end if
		else
			wheredetail = wheredetail & " and isNULL(p2.tplcompanyid,'')=''"
		end if

		if FRectchknotcash="Y" then
			wheredetail = wheredetail + " and m.ipkumdiv >=2 "
		else
        	wheredetail = wheredetail + " and m.ipkumdiv >3 "
		end if

		Select Case FRectIsPlusSaleItem
			Case "P"
				wheredetail = wheredetail & " and d.plus_sale_item_idx is not null "
			Case "N"
				wheredetail = wheredetail & " and d.plus_sale_item_idx is null "
		end Select

		IF (FRectIsSendGift="Y") THEN
			wheredetail = wheredetail & " and Exists(select f.orderserial from db_order.dbo.tbl_order_gift_data as f where f.orderserial=m.orderserial) "
		END IF

		''총갯수
		sqlStr = "select count(d.idx) as cnt, sum(d.itemno) as sumitemno, sum(d.itemno*d.itemcost) as sumitemcost, sum(d.itemno*d.buycash) as sumbuycash "
		If Left(FRectRegStart,4) < 2014 Then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m WITH(NOLOCK)"
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d WITH(NOLOCK)"
		Else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m WITH(NOLOCK)"
			sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d WITH(NOLOCK)"
		End If
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"

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

		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p2 with (nolock)"
		sqlStr = sqlStr & " 	on m.sitename=p2.id "
'		If FRectDispCate <> "" Then
'			sqlStr = sqlStr + " INNER JOIN [db_item].[dbo].[tbl_item] as i WITH(NOLOCK) ON d.itemid = i.itemid and i.dispcate1 = '" & FRectDispCate & "' "
'		End If
		sqlStr = sqlStr + " where d.itemid not in (0,100)"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + wheredetail

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		FSumItemNo = rsget("sumitemno")
		FSumItemCost = rsget("sumitemcost")
		FSumBuyCash = rsget("sumbuycash")
		rsget.close
		If FTotalCount = 0 Then
			FSumItemNo = 0
			FSumItemCost = 0
			FSumBuyCash = 0
		End If

		''데이타.
		sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " m.orderserial, m.jumundiv, d.itemno, d.itemid, d.itemname, d.buycash, d.itemcost,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.ipkumdate, m.regdate, m.buyname, m.reqname ,d.idx as detailidx, d.makerid, d.cancelyn as detailcancelyn, d.beasongdate, d.isupchebeasong as deliverytype, d.orgitemcost, d.itemcostCouponNotApplied"
		sqlStr = sqlStr + " ,d.omwdiv, m.sitename "
		sqlStr = sqlStr + " ,d.dlvfinishdt, d.jungsanfixdate "
		If Left(FRectRegStart,4) < 2014 Then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m WITH(NOLOCK)"
			sqlStr = sqlStr + "     Join [db_log].[dbo].tbl_old_order_detail_2003 d WITH(NOLOCK)"
		Else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m WITH(NOLOCK)"
			sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d WITH(NOLOCK)"
		End If
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"

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

		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p2 with (nolock)"
		sqlStr = sqlStr & " 	on m.sitename=p2.id "
		sqlStr = sqlStr + " where d.itemid not in (0,100)"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by m.orderserial desc, d.makerid asc"
		rsget.PageSize = FPageSize

		''response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FOrderserial = rsget("orderserial")
				FMasterItemList(i).FItemid 	 = rsget("itemid")
				FMasterItemList(i).FItemname    = rsget("itemname")
				FMasterItemList(i).FItemoption     = rsget("itemoptionname")
				FMasterItemList(i).FItemcnt     = rsget("itemno")
				FMasterItemList(i).FBuyname    = rsget("buyname")
				FMasterItemList(i).FReqname    = rsget("reqname")
				FMasterItemList(i).FMasterCancel	 = rsget("cancelyn")
				FMasterItemList(i).FOrderDate  = rsget("regdate")
				FMasterItemList(i).FIpkumDate = rsget("ipkumdate")
				FMasterItemList(i).FCurrstate  = rsget("baljuok")
				FMasterItemList(i).Fdetailidx = rsget("detailidx")
				FMasterItemList(i).FMakerid = rsget("makerid")
				FMasterItemList(i).FCancelYn = rsget("detailcancelyn")
				FMasterItemList(i).FDeliveryType = rsget("deliverytype")
				FMasterItemList(i).FJumunDiv = rsget("jumundiv")
				FMasterItemList(i).FBuyCash = rsget("buycash")
				FMasterItemList(i).FSellCash = rsget("itemcost")
				FMasterItemList(i).FOrgitemCost = rsget("orgitemcost")
				FMasterItemList(i).FitemcostCouponNotApplied = rsget("itemcostCouponNotApplied")
				FMasterItemList(i).FUpcheBeasongDate = rsget("beasongdate")
				FMasterItemList(i).Fsitename = rsget("sitename")
				if IsNull(FMasterItemList(i).FUpcheBeasongDate) then
					FMasterItemList(i).FUpcheBeasongDate = ""
				end if
                FMasterItemList(i).Fomwdiv = rsget("omwdiv")

				FMasterItemList(i).Fdlvfinishdt		= rsget("dlvfinishdt")
				if IsNull(FMasterItemList(i).Fdlvfinishdt) then
					FMasterItemList(i).Fdlvfinishdt = ""
				end if
				FMasterItemList(i).Fjungsanfixdate	= rsget("jungsanfixdate")
				if IsNull(FMasterItemList(i).Fjungsanfixdate) then
					FMasterItemList(i).Fjungsanfixdate = ""
				end if
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	public Sub DesignerDateSellListByItem()
		dim sqlStr,wheredetail
		dim i
		wheredetail = ""

		if (FRectDesignerID <>"") then
			wheredetail = " and d.makerid='" & FRectDesignerID & "'"
		end if

		if (FRectItemid <>"") then
			wheredetail = " and d.itemid='" & FRectItemid & "'"
		end if

		if (FRectSiteName <>"") then
			wheredetail = " and m.sitename='" & FRectSiteName & "'"
		end if

		if (FRectDateType="ipkumil") then
			'// 결제일기준
			wheredetail = wheredetail + " and m.ipkumdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.ipkumdate < '" & FRectRegEnd & "'"
		elseif (FRectDateType="chulgoil") then
			'// 출고일기준(배송비제외)
			wheredetail = wheredetail + " and d.beasongdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and d.beasongdate < '" & FRectRegEnd & "'"
			wheredetail = wheredetail + " and d.itemid <> 0 "
		elseif (FRectDateType="jungsanil") then
			'// 정산일기준
			wheredetail = wheredetail + " and d.jungsanfixdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and d.jungsanfixdate < '" & FRectRegEnd & "'"
		else
			'// 주문일기준
			wheredetail = wheredetail + " and m.regdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.regdate < '" & FRectRegEnd & "'"
		end if

		if (FRectDeliverType="upche") then
			wheredetail = wheredetail + " and d.isupchebeasong='Y'"

		elseif (FRectDeliverType="ten") then
			wheredetail = wheredetail + " and d.isupchebeasong<>'Y'"
		else

		end if

		Select Case FRectIsPlusSaleItem
			Case "P"
				wheredetail = wheredetail & " and d.plus_sale_item_idx is not null "
			Case "N"
				wheredetail = wheredetail & " and d.plus_sale_item_idx is null "
		end Select

		''총갯수
		sqlStr = "select count(distinct d.itemid) as cnt, sum(d.itemno) as sumitemno, sum(d.itemno*d.itemcost) as sumitemcost, sum(d.itemno*d.buycash) as sumbuycash "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"

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

		sqlStr = sqlStr + " where d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= '4'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + wheredetail
		''response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		FSumItemNo = rsget("sumitemno")
		FSumItemCost = rsget("sumitemcost")
		FSumBuyCash = rsget("sumbuycash")
		rsget.close

		If FTotalCount = 0 Then
			FSumItemNo = 0
			FSumItemCost = 0
			FSumBuyCash = 0
		End If

		''데이타.
		sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " sum(d.itemno) as itemno, d.itemid, d.itemname, sum(d.buycash*d.itemno) as buycash, sum(d.itemcost*d.itemno) as itemcost,"
		sqlStr = sqlStr + " d.itemoptionname, '0' as baljuok, sum(d.orgitemcost*d.itemno) as orgitemcost, sum(d.itemcostCouponNotApplied*d.itemno) as itemcostCouponNotApplied,"
		sqlStr = sqlStr + " '' as ipkumdate, '' as regdate, '' as buyname, '' as reqname ,0 as detailidx, d.makerid, '' as beasongdate, d.isupchebeasong as deliverytype, d.cancelyn as detailcancelyn, m.cancelyn"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"

		IF (FRectBrandPurchaseType<>"") then
		    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " on d.makerid=p.id"

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

		sqlStr = sqlStr + " where d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= '4'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid, d.itemname, d.itemoptionname, d.makerid, d.isupchebeasong, m.cancelyn, d.cancelyn"
		sqlStr = sqlStr + " order by d.makerid asc, d.itemid desc, d.itemoptionname"
		rsget.PageSize = FPageSize

		''response.write sqlStr
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FItemid 	 = rsget("itemid")
				FMasterItemList(i).FItemname    = rsget("itemname")
				FMasterItemList(i).FItemoption     = rsget("itemoptionname")
				FMasterItemList(i).FItemcnt     = rsget("itemno")
				FMasterItemList(i).FBuyname    = rsget("buyname")
				FMasterItemList(i).FReqname    = rsget("reqname")
				FMasterItemList(i).FMasterCancel	 = rsget("cancelyn")
				FMasterItemList(i).FOrderDate  = rsget("regdate")
				FMasterItemList(i).FIpkumDate = rsget("ipkumdate")
				FMasterItemList(i).FCurrstate  = rsget("baljuok")
				FMasterItemList(i).Fdetailidx = rsget("detailidx")
				FMasterItemList(i).FMakerid = rsget("makerid")
				FMasterItemList(i).FCancelYn = rsget("detailcancelyn")
				FMasterItemList(i).FDeliveryType = rsget("deliverytype")
				FMasterItemList(i).FBuyCash = rsget("buycash")
				FMasterItemList(i).FSellCash = rsget("itemcost")
				FMasterItemList(i).FOrgitemCost = rsget("orgitemcost")
				FMasterItemList(i).FitemcostCouponNotApplied = rsget("itemcostCouponNotApplied")
				FMasterItemList(i).FUpcheBeasongDate = rsget("beasongdate")
				if IsNull(FMasterItemList(i).FUpcheBeasongDate) then
					FMasterItemList(i).FUpcheBeasongDate = ""
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub


        ''[CS]업체배송관리>>업체배송관리
	public Sub DesignerDateMiBaljuMiBeasongList()
		dim sqlStr
		dim i
        ''차후 상품별 발주일(출고기준일)로 변경

        ' sqlStr = " select T.*"
        ' sqlStr = sqlStr + " ,u.userdiv, u.socname_kor, u.catecode, cl.code_nm, p.company_name, p.deliver_hp"
        ' sqlStr = sqlStr + " from ("
		' sqlStr = sqlStr + "     select d.makerid "
		' sqlStr = sqlStr + "     ,sum(case when (d.currstate = '0') then 1 else 0 end) as mitongbocnt" ''미통보 전체
		' sqlStr = sqlStr + "     ,sum(case when (d.currstate = '2') then 1 else 0 end) as mibaljucnt"  ''미확인 전체
		' sqlStr = sqlStr + "     ,sum(case when (d.currstate = '3') then 1 else 0 end) as mibeasongcnt" ''미출고 전체
		' sqlStr = sqlStr + "     ,sum(case when ((d.currstate = '2') and datediff(d,m.baljudate,getdate())<2) then 1 else 0 end ) as P_ndaymibaljucnt"
		' sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())>=2) and (d.currstate = '2') then 1 else 0 end ) as ndaymibaljucnt"
		' sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())<4) and (d.currstate = '3') then 1 else 0 end ) as P_ndaymibeasongcnt"
		' sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())>=4) and (d.currstate = '3') then 1 else 0 end ) as ndaymibeasongcnt"
		' sqlStr = sqlStr + "     from "
		' sqlStr = sqlStr + "     [db_order].[dbo].tbl_order_master m"
		' sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
		' sqlStr = sqlStr + "     on  m.orderserial=d.orderserial"

		' If FRectDispCDL <> "" Or FRectDispCate <> "" Then
		' 	sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
		' 	sqlStr = sqlStr + "     on  d.itemid = i.itemid "

		' 	if FRectDispCDL<>"" then
		' 		sqlStr = sqlStr + " and i.cate_large='" + FRectDispCDL + "'"
		' 	end if

		' 	if FRectDispCDM<>"" then
		' 		sqlStr = sqlStr + " and i.cate_mid='" + FRectDispCDM + "'"
		' 	end if

		' 	if FRectDispCDS<>"" then
		' 		sqlStr = sqlStr + " and i.cate_small='" + FRectDispCDS + "'"
		' 	end if

		' 	if FRectDispCate<>"" then
		' 		if LEN(FRectDispCate)>3 then
		' 			sqlStr = sqlStr + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		' 		end if
		' 		sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		' 	end if
		' End If

		' sqlStr = sqlStr + "     where m.regdate >='" & FRectRegStart & "'"
		' sqlStr = sqlStr + "     and m.regdate <='" & FRectRegEnd & "'"
        ' sqlStr = sqlStr + "     and m.cancelyn='N'"
        ' sqlStr = sqlStr + "     and m.jumundiv<>'9'"
        ' sqlStr = sqlStr + "     and m.ipkumdiv>'3'"
        ' sqlStr = sqlStr + "     and m.ipkumdiv<'8'"                       ''출고완료 제외
        ' sqlStr = sqlStr + "     and d.itemid<>0"
        ' if FRectDesignerID<>"" then
		' 	sqlStr = sqlStr + "     and d.makerid='" + CStr(FRectDesignerID) + "'"
		' end if
        ' sqlStr = sqlStr + "     and d.isupchebeasong='Y'"
        ' sqlStr = sqlStr + "     and d.cancelyn<>'Y'"
		' sqlStr = sqlStr + "     and d.currstate<'7'"                      ''출고완료 제외 / 인덱스 안타게.

		' sqlStr = sqlStr + "     group by d.makerid"
		' sqlStr = sqlStr + " ) T"

		' if FRectCDL<>"" then
		'     sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c u on T.makerid=u.userid"
		'     sqlStr = sqlStr + "     and u.catecode='"&FRectCDL&"'"
		' else
		'     sqlStr = sqlStr + " join [db_user].[dbo].tbl_user_c u on T.makerid=u.userid"
		' end if
		' sqlStr = sqlStr + " left join db_item.dbo.tbl_cate_large cl on u.catecode=cl.code_large"

		' sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on T.makerid=p.id"


		' sqlStr = sqlStr + " order by T.ndaymibaljucnt desc, ndaymibeasongcnt desc"

        ''차후 상품별 발주일(출고기준일)로 변경
		sqlStr = ""
		sqlStr = sqlStr & " SELECT T.makerid "
		sqlStr = sqlStr & " ,sum(case when (T.currstate = '0') then 1 else 0 end) mitongbocnt "
		sqlStr = sqlStr & " ,sum(case when (T.currstate = '2') then 1 else 0 end) as mibaljucnt "
		sqlStr = sqlStr & " ,sum(case when (T.currstate = '3') then 1 else 0 end) as mibeasongcnt "
		sqlStr = sqlStr & " ,sum(case when (T.cc <2) and (T.currstate = '2') then 1 else 0 end ) as P_ndaymibaljucnt  "
		sqlStr = sqlStr & " ,sum(case when (T.cc >=2) and (T.currstate = '2') then 1 else 0 end ) as ndaymibaljucnt  "
		sqlStr = sqlStr & " ,sum(case when (T.cc <4) and (T.currstate = '3') then 1 else 0 end ) as P_ndaymibeasongcnt "
		sqlStr = sqlStr & " ,sum(case when (T.cc >=4) and (T.currstate = '3') then 1 else 0 end ) as ndaymibeasongcnt  "
		sqlStr = sqlStr & " ,u.userdiv, u.socname_kor, u.catecode, cl.code_nm, p.company_name, p.deliver_hp  "
		sqlStr = sqlStr & " FROM ( "
		sqlStr = sqlStr & " 	SELECT d.makerid, d.currstate, m.baljudate, d.idx "
		sqlStr = sqlStr & " 	,(SELECT db_order.[dbo].[UF_GetDPlusDateStr] ('U', convert(varchar(10), m.baljudate, 23), convert(varchar(10), getdate(), 23))) as cc "
		sqlStr = sqlStr & " 	from [db_order].[dbo].tbl_order_master m  "
		sqlStr = sqlStr & " 	Join [db_order].[dbo].tbl_order_detail d on m.orderserial=d.orderserial  "
		If FRectDispCDL <> "" Or FRectDispCate <> "" Then
			sqlStr = sqlStr & "     Join [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr & "     on  d.itemid = i.itemid "

			if FRectDispCDL<>"" then
				sqlStr = sqlStr & " and i.cate_large='" + FRectDispCDL + "'"
			end if

			if FRectDispCDM<>"" then
				sqlStr = sqlStr & " and i.cate_mid='" + FRectDispCDM + "'"
			end if

			if FRectDispCDS<>"" then
				sqlStr = sqlStr & " and i.cate_small='" + FRectDispCDS + "'"
			end if

			if FRectDispCate<>"" then
				if LEN(FRectDispCate)>3 then
					sqlStr = sqlStr & " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
				end if
				sqlStr = sqlStr & " and i.itemid in (SELECT itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
			end if
		End If
		sqlStr = sqlStr & "     where m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr & "     and m.regdate <='" & FRectRegEnd & "'"
		sqlStr = sqlStr & " 	and m.cancelyn='N' "
		sqlStr = sqlStr & " 	and m.jumundiv<>'9' "
		sqlStr = sqlStr & " 	and m.ipkumdiv>'3' "
		sqlStr = sqlStr & " 	and m.ipkumdiv<'8' "                      ''출고완료 제외
		sqlStr = sqlStr & " 	and d.itemid<>0 "
		sqlStr = sqlStr & " 	and d.isupchebeasong='Y' "
		sqlStr = sqlStr & " 	and d.cancelyn<>'Y' "
		sqlStr = sqlStr & " 	and d.currstate<'7' "                      ''출고완료 제외 / 인덱스 안타게.
		if FRectDesignerID <> "" then
			sqlStr = sqlStr & "     and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if
		sqlStr = sqlStr & " 	GROUP BY d.makerid, d.idx, d.currstate, m.baljudate "
		sqlStr = sqlStr & " ) T "

		if FRectCDL<>"" then
		    sqlStr = sqlStr & "     Join [db_user].[dbo].tbl_user_c u on T.makerid=u.userid"
		    sqlStr = sqlStr & "     and u.catecode='"&FRectCDL&"'"
		else
		    sqlStr = sqlStr & " join [db_user].[dbo].tbl_user_c u on T.makerid=u.userid"
		end if
		sqlStr = sqlStr & " left join db_item.dbo.tbl_cate_large cl on u.catecode=cl.code_large"

		sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner p on T.makerid=p.id"
		sqlStr = sqlStr & " GROUP BY T.makerid,u.userdiv, u.socname_kor, u.catecode, cl.code_nm, p.company_name, p.deliver_hp "
		sqlStr = sqlStr & " order by sum(case when (T.cc >=2) and (T.currstate = '2') then 1 else 0 end )desc, sum(case when (T.cc >=4) and (T.currstate = '3') then 1 else 0 end ) desc "
If (session("ssBctID")="kjy8517") Then
	rw sqlStr
End If
		'rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpCheSMSItem

				FMasterItemList(i).FMakerid         = rsget("makerid")

				FMasterItemList(i).FNDayMiBaljuCnt = rsget("ndaymibaljucnt")
				FMasterItemList(i).FP_NDayMiBaljuCnt = rsget("p_ndaymibaljucnt")
				FMasterItemList(i).FMiBalJuCount    = rsget("mibaljucnt")
				FMasterItemList(i).Fmitongbocnt     = rsget("mitongbocnt")
				FMasterItemList(i).FNDayMiBeasongCnt = rsget("ndaymibeasongcnt")
				FMasterItemList(i).FP_NDayMiBeasongCnt = rsget("p_ndaymibeasongcnt")
				FMasterItemList(i).FMiBeasongCount  = rsget("mibeasongcnt")

				FMasterItemList(i).FUserDiv       = rsget("userdiv")
				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))

				FMasterItemList(i).FCompanyName    = db2html(rsget("company_name"))
				FMasterItemList(i).FDeliverHp       = db2html(rsget("deliver_hp"))

                FMasterItemList(i).Fcatecode    = rsget("catecode")
                FMasterItemList(i).Fcatename    = db2html(rsget("code_nm"))

				'FMasterItemList(i).FLastSendMsgDay  = rsget("")
				'if IsNULL(FMasterItemList(i).FDeliverHp) or (FMasterItemList(i).FDeliverHp="") then
				'	FMasterItemList(i).FDeliverHp       = rsget("manager_hp")
				'end if
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	''[CS]배송관리>>미출고리스트_업배 용
	public Sub getUpcheMichulgoList(byval isALL)
	    dim sqlStr, addSql
		dim i
		dim tmpSql, tmpDate

		Dim stOrderSerial, edOrderserial
		stOrderSerial = Mid(Replace(CStr(FRectRegStart),"-",""),3,6) + "00000"
		edOrderserial = Mid(Replace(CStr(FRectRegEnd),"-",""),3,6) + "00000"


		'// ===================================================================
		'' baljudate => 상품 (주문통보일=기준일) 로 변경
		addSql = " from [db_order].[dbo].tbl_order_master m with (nolock)"
		addSql = addSql + "     Join [db_order].[dbo].tbl_order_detail d with (nolock)"
		addSql = addSql + "     on m.orderserial=d.orderserial"

		If FRectDispCate <> "" Then
			addSql = addSql + "     Join [db_item].[dbo].tbl_item ii with (nolock)"
			addSql = addSql + "     on  d.itemid = ii.itemid "

			if FRectDispCate<>"" then
				if LEN(FRectDispCate)>3 then
					addSql = addSql + " and ii.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
				end if
				addSql = addSql + " and ii.itemid in (select itemid from db_item.dbo.tbl_display_cate_item with (nolock) where catecode like '" + FRectDispCate + "%' and isDefault='y') "
			end if
		End If

		addSql = addSql + "     left join [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
		addSql = addSql + "     on d.orderserial=T.orderserial"
		addSql = addSql + "     and d.idx=T.detailidx"
		addSql = addSql + " left join db_cs.dbo.tbl_cs_brand_memo B with (nolock)"
		addSql = addSql + " on "
		addSql = addSql + " 	B.brandid = d.makerid "
		addSql = addSql + " left join db_cs.dbo.tbl_cs_item_memo I with (nolock)"
		addSql = addSql + " on "
		addSql = addSql + " 	I.itemid = d.itemid "

		if FRectCDL<>"" then
		    addSql = addSql + "     Join [db_user].[dbo].tbl_user_c c with (nolock)"
		    addSql = addSql + "     on d.makerid=c.userid"
		    addSql = addSql + "     and c.catecode='"&FRectCDL&"'"
		end if

		IF (isALL) then
		        FRectRegEnd = LEft(CStr(dateAdd("d",1,now())),10)
		        FRectRegStart = LEft(CStr(dateAdd("m",-2,now())),10)

		        addSql = addSql + " where m.regdate >= '" & FRectRegStart & "'"
				addSql = addSql + " and m.regdate < '" & FRectRegEnd & "'"
		ELSE
				addSql = addSql + " where m.regdate >= '" & FRectRegStart & "'"
				addSql = addSql + " and m.regdate < '" & FRectRegEnd & "'"
		END IF

		if (FRectDetailState="MOO") then
		    addSql = addSql + " and m.ipkumdiv ='2'"
		else
		    addSql = addSql + " and m.ipkumdiv < '8'"
            addSql = addSql + " and m.ipkumdiv > '3'"
        end if
        addSql = addSql + " and m.cancelyn = 'N'"
        addSql = addSql + " and m.jumundiv <> '9' and m.jumundiv <> '7' "

		if (FRectDesignerID <>"") then
			addSql = addSql + " and d.makerid='" & FRectDesignerID & "'"
		end if
		addSql = addSql + " and d.itemid<>0"
		if (FRectItemid<>"") then
		    addSql = addSql + " and d.itemid="&FRectItemid&""
		end if

		if (FRectSiteName<>"") then
			if FRectSiteName="extall" then
				addSql = addSql + " and m.sitename <> '10x10'"
			else
		    	addSql = addSql + " and m.sitename = '" & FRectSiteName & "'"
			end if
		end if

		if (FRectDetailState="NOT7") then
		    addSql = addSql + " and d.currstate<'7'"
		elseif (FRectDetailState="MOO") then
		    addSql = addSql + " and d.currstate='0'"
		elseif (FRectDetailState="UP2") then
		    addSql = addSql + " and d.currstate>'1'"
		elseif (FRectDetailState="UP2NOT7") then
		    addSql = addSql + " and d.currstate>'1'"
		    addSql = addSql + " and d.currstate<'7'"
		elseif (FRectDetailState<>"") then
		    addSql = addSql + " and d.currstate='" & FRectDetailState&"'"
		end if
        addSql = addSql + " and d.isupchebeasong='Y'"
        addSql = addSql + " and d.cancelyn <> 'Y'"

        if (FRectMisendReason<>"") then
            if (FRectMisendReason="00") then
                addSql = addSql + "     and IsNULL(T.code,'00')='" & FRectMisendReason & "'"
            else
                addSql = addSql + "     and IsNULL(T.code,'00')='" & FRectMisendReason & "'"
            end if
        end if

        if (FRectMisendState="N") then
            addSql = addSql + "     and T.state is NULL"
        elseif (FRectMisendState<>"") then
            addSql = addSql + "     and T.state='" & FRectMisendState & "'"
        end if

        if (FRectdplusOver<>"") then

            'tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', " & FRectdplusOver & " " & VbCRLF
			tmpSql = " exec [db_cs].[dbo].[usp_getTENUpcheMinusWorkday] '" & Left(now(), 10) & "', " & FRectdplusOver & ", 'U' " & VbCRLF
            rsget.CursorLocation = adUseClient
            rsget.Open tmpSql, dbget, adOpenForwardOnly
        	if Not rsget.Eof then
                tmpDate = rsget("minusworkday")
            end if
        	rsget.close

			'// 근무일수 기준 D+4 일
			''addSql = addSql + "     and datediff(d,m.baljudate,getdate())>=" & FRectdplusOver
			addSql = addSql + "     and datediff(d,m.baljudate,'" & tmpDate & "') >= 0 "

        end if

        if (FRectdplusLower<>"") then
            addSql = addSql + "     and datediff(d,m.baljudate,getdate())<=" & FRectdplusLower
        end if

		'// 출고예정일 이전 주문 제외(출고예정일 이후 또는 품절출고불가 전부 표시)
        if (FRectExInMayChulgoDay="Y") then
            addSql = addSql + "     and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
        end if

		'// 출고소요일 이전 주문 제외(상품출고소요일->브랜드출고소요일->입력없으면 전부 표시)
        if (FRectExInNeedChulgoDay="Y") then
            addSql = addSql + "     and not ( "
            addSql = addSql + "     	((T.ipgodate is null) and (IsNull(T.code, '00') <> '05') and (IsNull(I.beasongneedday, 0) <> 0) and (datediff(d, DateAdd(d, I.beasongneedday, m.baljudate), getdate()) <= 0)) "
            addSql = addSql + "     	or "
            addSql = addSql + "     	((T.ipgodate is null) and (IsNull(T.code, '00') <> '05') and (IsNull(I.beasongneedday, 0) = 0) and (IsNull(B.beasongneedday, 0) <> 0) and (datediff(d, DateAdd(d, B.beasongneedday, m.baljudate), getdate()) <= 0)) "
            addSql = addSql + "     ) "
        end if

		'// 품절출고불가 제외
		if (FRectExStockOut = "Y") then
            addSql = addSql + " 	and IsNULL(T.code,'00') <> '05' "
		end if

		if (FRectExInNeedChulgoDay="Y") then

			'// TODO : 재입고예정일

		end if


		'// ===================================================================
        sqlStr = "select count(*) as cnt "

        sqlStr = sqlStr + addSql

		''rw     sqlStr & "<br><br>"

		IF (Not isALL) then
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			    FTotalCount = rsget("cnt")
			rsget.Close
		end IF


		'// ===================================================================
		sqlStr = "select top "&FPageSize*FCurrPage&" m.orderserial, d.itemno, d.itemid, d.itemname"
		sqlStr = sqlStr + " ,d.itemoptionname, isNull(d.currstate,0) as detailstate, d.upcheconfirmdate, d.beasongdate"
		sqlStr = sqlStr + " ,m.cancelyn, m.regdate, m.buyname, m.reqname , d.makerid, d.idx as detailidx "
		sqlStr = sqlStr + " ,m.baljudate, T.code, T.state, T.ipgodate, T.regdate as misendregdate, m.sitename "
		sqlStr = sqlStr + " , ( "
		sqlStr = sqlStr + " 	select count(*) as csMemoCnt "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	db_cs.dbo.tbl_cs_memo cm "
		sqlStr = sqlStr + " 	where cm.orderserial = m.orderserial "
		sqlStr = sqlStr + " ) as csMemoCnt "
		sqlStr = sqlStr + " , (select db_order.[dbo].[UF_GetDPlusDateStr] ('U', convert(varchar(10), m.baljudate, 23), convert(varchar(10), isnull(d.beasongdate, getdate()), 23))) as dday "
		sqlStr = sqlStr + addSql

		if (FRectSortBy = "makerid") then
			sqlStr = sqlStr + " order by d.makerid, d.itemid, d.itemoption"
		elseif (FRectSortBy = "orderserial") then
			sqlStr = sqlStr + " order by m.orderserial, d.itemid, d.itemoption"
		else
			sqlStr = sqlStr + " order by isNull(m.baljudate,getdate()+365),  d.currstate, d.makerid, m.orderserial, d.itemid, d.itemoption"
		end if

''		rw     sqlStr
		rsget.PageSize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		IF (isALL) then
		    FTotalCount = FResultCount
		END IF

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

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FOrderserial 	= rsget("orderserial")
				FMasterItemList(i).FItemid 	    	= rsget("itemid")
				FMasterItemList(i).FItemname    	= db2html(rsget("itemname"))
				FMasterItemList(i).FItemoption     	= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FItemcnt     	= rsget("itemno")
				FMasterItemList(i).FBuyname     	= db2html(rsget("buyname"))
				FMasterItemList(i).FReqname     	= db2html(rsget("reqname"))
				FMasterItemList(i).FCancelYn	 	= rsget("cancelyn")
				FMasterItemList(i).FRegdate     	= rsget("regdate")
				FMasterItemList(i).FCurrstate   	= rsget("detailstate")
				FMasterItemList(i).FMakerid     	= rsget("makerid")

                FMasterItemList(i).Fbaljudate   	= rsget("baljudate")
                FMasterItemList(i).FUpcheConfirmDate = rsget("upcheconfirmdate")
                FMasterItemList(i).FUpcheBeasongDate = rsget("beasongdate")

                FMasterItemList(i).FMisendReason  	= rsget("code")
                FMasterItemList(i).FMisendState   	= rsget("state")
                FMasterItemList(i).FMisendipgodate	= rsget("ipgodate")

                FMasterItemList(i).Fmisendregdate 	= rsget("misendregdate")

                FMasterItemList(i).Fdetailidx 		= rsget("detailidx")

				FMasterItemList(i).Fsitename 		= rsget("sitename")

				FMasterItemList(i).FcsMemoCnt 		= rsget("csMemoCnt")
				FMasterItemList(i).FDday 		= rsget("dday")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close

    end Sub

	' /admin/upchebeasong/upchemibeasonglistNEW.asp
	public Sub getUpcheMichulgoListNEW(byval isALL)
	    dim sqlStr, addSql, i

		If (FRectDeliverType = "") Then
			FRectDeliverType = "Y"
		End If

		If (FRectItemid = "") Then
			FRectItemid = 0
		End If

		If (FRectdplusOver = "") Then
			FRectdplusOver = 0
		End If

		If (FRectdplusLower = "") Then
			FRectdplusLower = 0
		End If

		If (isALL = True) Then
			sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_CS_Cnt '" & FRectDeliverType & "', '" & FRectDesignerID & "', " & FRectItemid & ", '" & FRectSiteName & "', " & FRectdplusOver & ", " & FRectdplusLower & ", '" & FRectMisendReason & "', '" & FRectMisendState & "', '" & FRectExInMayChulgoDay & "', '" & FRectExStockOut & "', '" & FRectUpcheNoCheck & "', '" & FRectExToday & "', '"& frectdetailcancelyn &"', '" & FRectIncIpkumdiv4 & "', '" & FRectItemOption & "' "

			response.write sqlStr & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FOrderCnt = rsget("orderCnt")
			FSumItemNo = rsget("totItemNo")
			rsget.Close
		End If

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_CS_List " & FPageSize & ", " & FCurrPage & ", '" & FRectDeliverType & "', '" & FRectDesignerID & "', " & FRectItemid & ", '" & FRectSiteName & "', " & FRectdplusOver & ", " & FRectdplusLower & ", '" & FRectMisendReason & "', '" & FRectMisendState & "', '" & FRectExInMayChulgoDay & "', '" & FRectExStockOut & "', '" & FRectUpcheNoCheck & "', '" & FRectExToday & "', '"& frectdetailcancelyn &"', '" & FRectIncIpkumdiv4 & "', '" & FRectItemOption & "' "

		response.write sqlStr & "<br>"
        ''response.end
		rsget.PageSize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		If (isALL <> True) Then
			FTotalCount = FResultCount
		End If

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

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FOrderserial 	= rsget("orderserial")
				FMasterItemList(i).FItemid 	    	= rsget("itemid")
				FMasterItemList(i).FItemname    	= db2html(rsget("itemname"))
				FMasterItemList(i).FItemoption     	= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FItemcnt     	= rsget("itemno")
				FMasterItemList(i).FBuyname     	= db2html(rsget("buyname"))
				FMasterItemList(i).FReqname     	= db2html(rsget("reqname"))
				FMasterItemList(i).FCancelYn	 	= rsget("cancelyn")
				FMasterItemList(i).FRegdate     	= rsget("regdate")
				FMasterItemList(i).FCurrstate   	= rsget("detailstate")
				FMasterItemList(i).FMakerid     	= rsget("makerid")

                FMasterItemList(i).Fbaljudate   	= rsget("baljudate")
                FMasterItemList(i).Fipkumdate   	= rsget("ipkumdate")
                FMasterItemList(i).FUpcheConfirmDate = rsget("upcheconfirmdate")
                FMasterItemList(i).FUpcheBeasongDate = rsget("beasongdate")

                FMasterItemList(i).FMisendReason  	= rsget("code")
                FMasterItemList(i).FMisendState   	= rsget("state")
                FMasterItemList(i).FMisendipgodate	= rsget("ipgodate")

                FMasterItemList(i).Fmisendregdate 	= rsget("misendregdate")
                FMasterItemList(i).Fdetailidx 		= rsget("detailidx")
				FMasterItemList(i).Fsitename 		= rsget("sitename")
				FMasterItemList(i).FcsMemoCnt 		= rsget("csMemoCnt")

				FMasterItemList(i).Fmisendreguserid 	= rsget("misendreguserid")

				FMasterItemList(i).Fmisendmodiuserid 	= rsget("misendmodiuserid")
				FMasterItemList(i).Fmisendmodidate 		= rsget("misendmodidate")
				FMasterItemList(i).FsendCount 			= rsget("sendCount")
				FMasterItemList(i).FlastSendUserid 		= rsget("lastSendUserid")
				FMasterItemList(i).FlastSendDate 		= rsget("lastSendDate")
				FMasterItemList(i).FDetailCancelYn 		= rsget("DetailCancelYn")
				FMasterItemList(i).FDday				= rsget("dday")
                FMasterItemList(i).FDdayByIpkumdate		= rsget("ddayByIpkumdate")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close

    end Sub

	public Sub getUpcheMichulgoListByBrand()
	    dim sqlStr, addSql
		dim i

		If (FRectDeliverType = "") Then
			FRectDeliverType = "Y"
		End If

		If (FRectItemid = "") Then
			FRectItemid = 0
		End If

		If (FRectdplusOver = "") Then
			FRectdplusOver = 0
		End If

		If (FRectdplusLower = "") Then
			FRectdplusLower = 0
		End If

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_CS_List_Brand " & FPageSize & ", " & FCurrPage & ", '" & FRectDeliverType & "', '" & FRectDesignerID & "', " & FRectItemid & ", '" & FRectSiteName & "', " & FRectdplusOver & ", " & FRectdplusLower & ", '" & FRectMisendReason & "', '" & FRectMisendState & "', '" & FRectExInMayChulgoDay & "', '" & FRectExStockOut & "', '" & FRectUpcheNoCheck & "', '" & FRectExToday & "', '" & FRectIncIpkumdiv4 & "' "

		rw     sqlStr
		''response.end
		rsget.PageSize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		FTotalCount = FResultCount

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

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FMakerid     	= rsget("makerid")
				FMasterItemList(i).FItemcnt     	= rsget("cnt")
                FMasterItemList(i).Fvacation     	= rsget("vacation")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close

    end Sub

	''[CS]업체배송관리>>업체배송목록 /팝업
    public Sub getUpchebeasongList()
        dim sqlStr
		dim i
        '' baljudate => 상품 (주문통보일=기준일) 로 변경
        sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_mibeasong_list T "
		sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		sqlStr = sqlStr + "     and d.idx=T.detailidx"
		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
		    sqlStr = sqlStr + "     on d.makerid=c.userid"
		    sqlStr = sqlStr + "     and c.catecode='"&FRectCDL&"'"
		end if
		sqlStr = sqlStr + " where m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		if (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and m.ipkumdiv =2"
		else
            sqlStr = sqlStr + " and m.ipkumdiv > 3"
            sqlStr = sqlStr + " and m.ipkumdiv < 8"
        end if
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv <> '9'"
        if (FRectDesignerID <>"") then
			sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		end if
		sqlStr = sqlStr + " and d.itemid<>0"

		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid&""
		end if

		if (FRectDetailState="NOT7") then
		    sqlStr = sqlStr + " and d.currstate<7"
		elseif (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and d.currstate=0"
		elseif (FRectDetailState="UP2") then
		    sqlStr = sqlStr + " and d.currstate>1"
		elseif (FRectDetailState="UP2NOT7") then
		    sqlStr = sqlStr + " and d.currstate>1"
		    sqlStr = sqlStr + " and d.currstate<7"
		elseif (FRectDetailState<>"") then
		    sqlStr = sqlStr + " and d.currstate=" & FRectDetailState
		end if
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"

        if (FRectMisendReason<>"") then
            if (FRectMisendReason="00") then
                sqlStr = sqlStr + "     and IsNULL(T.code,'00')='" & FRectMisendReason & "'"
            else
                sqlStr = sqlStr + "     and T.code='" & FRectMisendReason & "'"
            end if
        end if

        if (FRectMisendState<>"") then
            sqlStr = sqlStr + "     and T.state='" & FRectMisendState & "'"
        end if

        if (FRectdplusOver<>"") then
            sqlStr = sqlStr + "     and datediff(d,m.baljudate,getdate())>=" & FRectdplusOver
        end if
''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close



		sqlStr = "select top "&FPageSize*FCurrPage&" m.orderserial, d.itemno, d.itemid, d.itemname"
		sqlStr = sqlStr + " ,d.itemoptionname, isNull(d.currstate,0) as detailstate, d.upcheconfirmdate, d.beasongdate"
		sqlStr = sqlStr + " ,m.cancelyn, m.regdate, m.buyname, m.reqname , d.makerid"
		sqlStr = sqlStr + " ,m.baljudate, T.code, T.state, T.ipgodate, T.regdate as misendregdate "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_mibeasong_list T "
		sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		sqlStr = sqlStr + "     and d.idx=T.detailidx"
		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
		    sqlStr = sqlStr + "     on d.makerid=c.userid"
		    sqlStr = sqlStr + "     and c.catecode='"&FRectCDL&"'"
		end if
		sqlStr = sqlStr + " where m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		if (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and m.ipkumdiv =2"
		else
            sqlStr = sqlStr + " and m.ipkumdiv > 3"
            sqlStr = sqlStr + " and m.ipkumdiv < 8"
        end if
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv <> '9'"



		if (FRectDesignerID <>"") then
			sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		end if
		sqlStr = sqlStr + " and d.itemid<>0"
		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid&""
		end if
		if (FRectDetailState="NOT7") then
		    sqlStr = sqlStr + " and d.currstate<7"
		elseif (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and d.currstate=0"
		elseif (FRectDetailState="UP2") then
		    sqlStr = sqlStr + " and d.currstate>1"
		elseif (FRectDetailState="UP2NOT7") then
		    sqlStr = sqlStr + " and d.currstate>1"
		    sqlStr = sqlStr + " and d.currstate<7"
		elseif (FRectDetailState<>"") then
		    sqlStr = sqlStr + " and d.currstate=" & FRectDetailState
		end if
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"


        if (FRectMisendReason<>"") then
            if (FRectMisendReason="00") then
                sqlStr = sqlStr + "     and IsNULL(T.code,'00')='" & FRectMisendReason & "'"
            else
                sqlStr = sqlStr + "     and T.code='" & FRectMisendReason & "'"
            end if
        end if

        if (FRectMisendState<>"") then
            sqlStr = sqlStr + "     and T.state='" & FRectMisendState & "'"
        end if

        if (FRectdplusOver<>"") then
            sqlStr = sqlStr + "     and datediff(d,m.baljudate,getdate())>=" & FRectdplusOver
        end if

		sqlStr = sqlStr + " order by isNull(m.baljudate,getdate()+3650),  IsNULL(d.currstate,0), d.idx "

		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

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

				set FMasterItemList(i) = new CUpchebeasongDetail

				FMasterItemList(i).FOrderserial = rsget("orderserial")
				FMasterItemList(i).FItemid 	    = rsget("itemid")
				FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
				FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
				FMasterItemList(i).FItemcnt     = rsget("itemno")
				FMasterItemList(i).FBuyname     = db2html(rsget("buyname"))
				FMasterItemList(i).FReqname     = db2html(rsget("reqname"))
				FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
				FMasterItemList(i).FRegdate     = rsget("regdate")
				FMasterItemList(i).FCurrstate   = rsget("detailstate")
				FMasterItemList(i).FMakerid     = rsget("makerid")

                FMasterItemList(i).Fbaljudate   = rsget("baljudate")
                FMasterItemList(i).FUpcheConfirmDate = rsget("upcheconfirmdate")
                FMasterItemList(i).FUpcheBeasongDate = rsget("beasongdate")

                FMasterItemList(i).FMisendReason  = rsget("code")
                FMasterItemList(i).FMisendState   = rsget("state")
                FMasterItemList(i).FMisendipgodate= rsget("ipgodate")

                FMasterItemList(i).Fmisendregdate = rsget("misendregdate")
				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close
    end Sub

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

' 사용중지. 디비에서 받아옴 2019.09.17 한용민
public function GetMichulgoSMSString(misendReason)
	select Case misendReason
		'// 출고지연
		CASE "03" : GetMichulgoSMSString = "[텐바이텐 출고지연안내]주문하신 상품 중 [상품명]([상품코드]) 상품이 [출고예정일]에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."

		'// 주문제작
		CASE "02" : GetMichulgoSMSString = "[텐바이텐 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 주문제작 상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 수입
		CASE "08" : GetMichulgoSMSString = "[텐바이텐 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 수입상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 가구배송
		CASE "09" : GetMichulgoSMSString = "[텐바이텐 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 가구상품으로 [출고예정일]에 발송될 예정이며, 우천시 변경될 수 있음 양해 부탁드립니다. 직배송 상품으로 배송 전 별도 연락드릴 예정입니다. 감사합니다."

		'// 예약배송
		CASE "04" : GetMichulgoSMSString = "[텐바이텐 출고예정안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 예약배송 상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 업체휴가
		CASE "10" : GetMichulgoSMSString = "[텐바이텐 출고지연안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 업체 휴가로 인해 [출고예정일]에 발송될 예정입니다. 빠른 배송 드리지 못해 죄송합니다."

		'// 고객지정배송
		CASE "07" : GetMichulgoSMSString = "[텐바이텐 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 고객지정배송상품으로 [출고예정일]에 발송될 예정입니다. 감사합니다."

		CASE ELSE : GetMichulgoSMSString = ""
	end Select
end function

' 사용중지. 디비에서 받아옴 2019.09.17 한용민
public function GetMichulgoMailString(misendReason)
	dim mailText

	mailText = ""
	select Case misendReason
		'// 출고지연
		CASE "03" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품이 발송이 지연될 예정입니다.\n"
			mailText = mailText + "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.\n"
			mailText = mailText + "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,\n"
			mailText = mailText + "고객행복센터로 연락 부탁 드립니다.\n"
			mailText = mailText + "쇼핑에 불편을 드린 점 진심으로 사과 드리며, 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.\n"

		'// 주문제작
		CASE "02" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품은 주문 후 제작 되는 상품으로\n"
			mailText = mailText + "일반상품과 달리 주문제작에 기간이 소요되는 상품입니다.\n"
			mailText = mailText + "아래와 같이 발송예정일을 안내해드리오니,\n"
			mailText = mailText + "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.\n"

		'// 수입
		CASE "08" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품은 제품 수입 후 발송되는 상품으로\n"
			mailText = mailText + "일반상품과 달리 상품 수입에 조금 더 기간이 소요되는 상품입니다.\n"
			mailText = mailText + "아래와 같이 발송예정일을 안내해드리오니,\n"
			mailText = mailText + "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.\n"

		'// 가구배송
		CASE "09" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품은 가구 상품으로\n"
			mailText = mailText + "일반상품과 달리 배송에 조금 더 기간이 소요되는 상품입니다.\n"
			mailText = mailText + "아래와 같이 발송예정일을 안내해드리오니,\n"
			mailText = mailText + "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.\n"
			mailText = mailText + "가구 배송상품으로 우천시 일정이 조금 변경될 수 있으며, \n"
			mailText = mailText + "직배송 상품으로 배송 전 별도 연락 드리겠습니다.\n"
			mailText = mailText + "이와 관련, 문의사항은 고객행복센터로 연락 부탁드립니다.\n"

		'// 예약배송
		CASE "04" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품의 출고안내메일입니다.\n"
			mailText = mailText + "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,\n"
			mailText = mailText + "부득이한 사정으로 상품취소를 원하시는 경우,\n"
			mailText = mailText + "고객행복센터로 연락 부탁드립니다.\n"

		'// 업체휴가
		CASE "10" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품이 업체 휴가 기간으로 인해 발송이 지연될 예정입니다.\n"
			mailText = mailText + "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.\n"
			mailText = mailText + "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,\n"
			mailText = mailText + "고객행복센터로 연락 부탁 드립니다.\n"
			mailText = mailText + "쇼핑에 불편을 드린 점 진심으로 사과 드리며, 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.\n"

		'// 고객지정배송
		CASE "07" :
			mailText = mailText + "안녕하세요. 고객님\n\n"
			mailText = mailText + "고객님께서 주문하신 상품의 출고안내 메일입니다.\n"
			mailText = mailText + "주문하신 상품은 고객지정배송상품으로 아래 발송예정일에 발송될 예정이며,\n"
			mailText = mailText + "부득이한 사정으로 상품취소를 원하시는 경우,\n"
			mailText = mailText + "고객행복센터로 연락 부탁드립니다.\n"

		CASE ELSE :
			mailText = ""

	end Select

	GetMichulgoMailString = mailText
end function

public function GetMichulgoMailTitleString(misendReason)
	select Case misendReason
		'// 출고지연
		CASE "03" : GetMichulgoMailTitleString = "[텐바이텐] 출고지연안내 메일입니다."

		'// 주문제작
		CASE "02" : GetMichulgoMailTitleString = "[텐바이텐] 출고 일정 안내 메일입니다."

		'// 수입
		CASE "08" : GetMichulgoMailTitleString = "[텐바이텐] 출고 일정 안내 메일입니다."

		'// 가구배송
		CASE "09" : GetMichulgoMailTitleString = "[텐바이텐] 출고 일정 안내 메일입니다."

		'// 예약배송
		CASE "04" : GetMichulgoMailTitleString = "[텐바이텐] 출고예정안내 메일입니다."

		'// 업체휴가
		CASE "10" : GetMichulgoMailTitleString = "[텐바이텐] 출고지연안내 메일입니다."

		'// 고객지정배송
		CASE "07" : GetMichulgoMailTitleString = "[텐바이텐] 출고 일정 안내 메일입니다."

		CASE ELSE : GetMichulgoMailTitleString = ""
	end Select
end function

%>
