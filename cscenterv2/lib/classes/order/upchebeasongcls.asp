<%
Class CUpCheSMSItem
	public FMakerid
	public FCompanyName
	public Fmitongbocnt
	public FMiBalJuCount
	public FMiBeasongCount
	public FLastSendMsgDay
	public FDeliverHp
	public FDeliverPhone
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
            CASE "00" : getMisendText = "입력대기"
            CASE "01" : getMisendText = "재고부족"
            CASE "04" : getMisendText = "예약상품"

            CASE "02" : getMisendText = "주문제작"
            CASE "52" : getMisendText = "주문제작"
            CASE "03" : getMisendText = "출고지연"
            CASE "53" : getMisendText = "출고지연"
            CASE "05" : getMisendText = "품절출고불가"
            CASE "55" : getMisendText = "품절출고불가"
            CASE ELSE : getMisendText = FMisendReason
        end Select
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
    public FRectIpkumdiv

	public FRectDateType
	public FRectDeliverType
    public FRect

    public FRectCDL
    public FRectDetailState
    public FRectMisendReason
    public FRectMisendState
    public FRectdplusOver

    public FRectCurrState	' 상태

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
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,"
		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d, "

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
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m, " & TABLE_ORDERDETAIL & " d"
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
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m, " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + " (select distinct top 500 m.orderserial "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m, " & TABLE_ORDERDETAIL & " d"
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
			wheredetail = " and d.makerid='" & FRectDesignerID & "'"
		end if

		if (FRectItemid <>"") then
			wheredetail = " and d.itemid='" & FRectItemid & "'"
		end if

		if (FRectDateType="ipkumil") then
			wheredetail = wheredetail + " and m.ipkumdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.ipkumdate < '" & FRectRegEnd & "'"
		else
			wheredetail = wheredetail + " and m.regdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.regdate < '" & FRectRegEnd & "'"
		end if

		if (FRectDispCate<>"") then
			wheredetail = wheredetail + " and exists(Select 1 from db_item.dbo.tbl_display_cate_item as c where c.isDefault='y' " &_
									" and c.itemid=d.itemid " &_
									" and c.catecode like '" & FRectDispCate & "%')"
		end if

		if (FRectDeliverType="upche") then
			wheredetail = wheredetail + " and d.isupchebeasong='Y'"

		elseif (FRectDeliverType="ten") then
			wheredetail = wheredetail + " and d.isupchebeasong<>'Y'"
		else

		end if
		''#################################################
		''총갯수
		''#################################################
		sqlStr = "select count(d.idx) as cnt "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= '4'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + wheredetail

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.close
		''#################################################
		''데이타.
		''#################################################

		sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " m.orderserial, m.jumundiv, d.itemno, d.itemid, d.itemname, d.buycash, d.itemcost,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.ipkumdate, m.regdate, m.buyname, m.reqname ,d.idx as detailidx, d.makerid, d.cancelyn as detailcancelyn, d.beasongdate, d.isupchebeasong as deliverytype"
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m "
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= '4'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
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

        sqlStr = " select T.*"
        sqlStr = sqlStr + " ,u.userdiv, u.socname_kor, u.catecode, cl.code_nm, p.company_name, p.deliver_hp, p.deliver_phone"
        sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + "     select d.makerid "
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '0') then 1 else 0 end) as mitongbocnt" ''미통보 전체
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '2') then 1 else 0 end) as mibaljucnt"  ''미확인 전체
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '3') then 1 else 0 end) as mibeasongcnt" ''미출고 전체
		sqlStr = sqlStr + "     ,sum(case when ((d.currstate = '2') and datediff(d,m.baljudate,getdate())<2) then 1 else 0 end ) as P_ndaymibaljucnt"
		sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())>=2) and (d.currstate = '2') then 1 else 0 end ) as ndaymibaljucnt"
		sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())<4) and (d.currstate = '3') then 1 else 0 end ) as P_ndaymibeasongcnt"
		sqlStr = sqlStr + "     ,sum(case when (datediff(d,m.baljudate,getdate())>=4) and (d.currstate = '3') then 1 else 0 end ) as ndaymibeasongcnt"
		sqlStr = sqlStr + "     from "
		sqlStr = sqlStr + "     " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on  m.orderserial=d.orderserial"

		sqlStr = sqlStr + "     where m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr + "     and m.regdate <='" & FRectRegEnd & "'"
        sqlStr = sqlStr + "     and m.cancelyn='N'"
        sqlStr = sqlStr + "     and m.jumundiv<>'9'"
        sqlStr = sqlStr + "     and m.ipkumdiv>'3'"
        sqlStr = sqlStr + "     and m.ipkumdiv<'8'"                       ''출고완료 제외
        sqlStr = sqlStr + "     and d.itemid<>0"
        if FRectDesignerID<>"" then
			sqlStr = sqlStr + "     and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if
        sqlStr = sqlStr + "     and d.isupchebeasong='Y'"
        sqlStr = sqlStr + "     and d.cancelyn<>'Y'"
		sqlStr = sqlStr + "     and d.currstate<'7'"                      ''출고완료 제외 / 인덱스 안타게.

		sqlStr = sqlStr + "     group by d.makerid"
		sqlStr = sqlStr + " ) T"

		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join " & TABLE_USER_C & " u on T.makerid=u.userid"
		    sqlStr = sqlStr + "     and u.catecode='"&FRectCDL&"'"
		else
		    sqlStr = sqlStr + " join " & TABLE_USER_C & " u on T.makerid=u.userid"
		end if
		sqlStr = sqlStr + " left join " & TABLE_CATEGORY_LARGE & " cl on u.catecode=cl.code_large"

		sqlStr = sqlStr + " left join " & TABLE_PARTNER & " p on T.makerid=p.id"


		sqlStr = sqlStr + " order by T.ndaymibaljucnt desc, ndaymibeasongcnt desc, T.makerid"
'rw sqlStr
		rsget.Open sqlStr,dbget,1

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
                FMasterItemList(i).FDeliverPhone       = db2html(rsget("deliver_phone"))

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
	    dim sqlStr
		dim i

		Dim stOrderSerial, edOrderserial
		stOrderSerial = Mid(Replace(CStr(FRectRegStart),"-",""),3,6) + "00000"
		edOrderserial = Mid(Replace(CStr(FRectRegEnd),"-",""),3,6) + "00000"

        '' baljudate => 상품 (주문통보일=기준일) 로 변경
        sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		if (FRectMisendReason<>"") or (FRectMisendState="N") then
		    sqlStr = sqlStr + "     left join " & TABLE_MIBEASONG_LIST & " T "
		    sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		    sqlStr = sqlStr + "     and d.detailidx=T.detailidx"
    	elseif (FRectMisendState<>"") then
    	    sqlStr = sqlStr + "     join " & TABLE_MIBEASONG_LIST & " T "
		    sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		    sqlStr = sqlStr + "     and d.detailidx=T.detailidx"
    	end if

		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join " & TABLE_USER_C & " c"
		    sqlStr = sqlStr + "     on d.makerid=c.userid"
		    sqlStr = sqlStr + "     and c.catecode='"&FRectCDL&"'"
		end if
		sqlStr = sqlStr + " where m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		if (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and m.ipkumdiv ='2'"
		else
		    sqlStr = sqlStr + " and m.ipkumdiv < '8'"
            sqlStr = sqlStr + " and m.ipkumdiv > '3'"
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
		    sqlStr = sqlStr + " and d.currstate<'7'"
		elseif (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and d.currstate='0'"
		elseif (FRectDetailState="UP2") then
		    sqlStr = sqlStr + " and d.currstate>'1'"
		elseif (FRectDetailState="UP2NOT7") then
		    sqlStr = sqlStr + " and d.currstate>'1'"
		    sqlStr = sqlStr + " and d.currstate<'7'"
		elseif (FRectDetailState<>"") then
		    sqlStr = sqlStr + " and d.currstate='" & FRectDetailState &"'"
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

        if (FRectMisendState="N") then
            sqlStr = sqlStr + "     and T.state is NULL"
        elseif (FRectMisendState<>"") then
            sqlStr = sqlStr + "     and T.state='" & FRectMisendState & "'"
        end if

        if (FRectdplusOver<>"") then
            sqlStr = sqlStr + "     and datediff(d,m.baljudate,getdate())>=" & FRectdplusOver
        end if

IF (Not isALL) then
''rw     sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close
end IF


		sqlStr = "select top "&FPageSize*FCurrPage&" m.orderserial, d.itemno, d.itemid, d.itemname"
		sqlStr = sqlStr + " ,d.itemoptionname, isNull(d.currstate,0) as detailstate, d.upcheconfirmdate, d.beasongdate"
		sqlStr = sqlStr + " ,m.cancelyn, m.regdate, m.buyname, m.reqname , d.makerid"
		sqlStr = sqlStr + " ,m.baljudate, T.code, T.state, T.ipgodate, T.regdate as misendregdate "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		if (FRectMisendState<>"") and (FRectMisendState<>"N") then
    	    sqlStr = sqlStr + "     join " & TABLE_MIBEASONG_LIST & " T "
		    sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		    sqlStr = sqlStr + "     and d.detailidx=T.detailidx"
		else
    		sqlStr = sqlStr + "     left join " & TABLE_MIBEASONG_LIST & " T "
    		sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
    		sqlStr = sqlStr + "     and d.detailidx=T.detailidx"
    	end if

		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join " & TABLE_USER_C & " c"
		    sqlStr = sqlStr + "     on d.makerid=c.userid"
		    sqlStr = sqlStr + "     and c.catecode='"&FRectCDL&"'"
		end if
IF (isALL) then
        FRectRegEnd = LEft(CStr(dateAdd("d",1,now())),10)
        FRectRegStart = LEft(CStr(dateAdd("m",-2,now())),10)

        sqlStr = sqlStr + " where m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
ELSE
		sqlStr = sqlStr + " where m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
END IF
		if (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and m.ipkumdiv ='2'"
		else
		    sqlStr = sqlStr + " and m.ipkumdiv < '8'"
            sqlStr = sqlStr + " and m.ipkumdiv > '3'"
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
		    sqlStr = sqlStr + " and d.currstate<'7'"
		elseif (FRectDetailState="MOO") then
		    sqlStr = sqlStr + " and d.currstate='0'"
		elseif (FRectDetailState="UP2") then
		    sqlStr = sqlStr + " and d.currstate>'1'"
		elseif (FRectDetailState="UP2NOT7") then
		    sqlStr = sqlStr + " and d.currstate>'1'"
		    sqlStr = sqlStr + " and d.currstate<'7'"
		elseif (FRectDetailState<>"") then
		    sqlStr = sqlStr + " and d.currstate='" & FRectDetailState&"'"
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

        if (FRectMisendState="N") then
            sqlStr = sqlStr + "     and T.state is NULL"
        elseif (FRectMisendState<>"") then
            sqlStr = sqlStr + "     and T.state='" & FRectMisendState & "'"
        end if

        if (FRectdplusOver<>"") then
            sqlStr = sqlStr + "     and datediff(d,m.baljudate,getdate())>=" & FRectdplusOver
        end if

		sqlStr = sqlStr + " order by isNull(m.baljudate,getdate()+365),  d.currstate"

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

	''[CS]업체배송관리>>업체배송목록 /팝업
    public Sub getUpchebeasongList()
        dim sqlStr
		dim i
        '' baljudate => 상품 (주문통보일=기준일) 로 변경
        sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     left join " & TABLE_MIBEASONG_LIST & " T "
		sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		sqlStr = sqlStr + "     and d.idx=T.detailidx"
		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join " & TABLE_USER_C & " c"
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
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     Join " & TABLE_ORDERDETAIL & " d"
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     left join " & TABLE_MIBEASONG_LIST & " T "
		sqlStr = sqlStr + "     on d.orderserial=T.orderserial"
		sqlStr = sqlStr + "     and d.idx=T.detailidx"
		if FRectCDL<>"" then
		    sqlStr = sqlStr + "     Join " & TABLE_USER_C & " c"
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

'    public Sub DesignerDateBaljuList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuCount()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub UpchebeasongMibaljuList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuDetail()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongCount()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongNdayList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuNdayList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongDetailList()
'        response.write "사용중지 - 관리자 문의 요망"
'        dbget.close()	:	response.End
'    end Sub


    ''''Maybe NotUsing..
''	public Sub DesignerDateBaljuList()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select m.orderserial, d.itemno, d.itemid, d.itemname,"
''		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
''		sqlStr = sqlStr + " m.cancelyn, m.regdate, m.buyname, m.reqname , d.makerid"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,  " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'' 		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
''        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
''        sqlStr = sqlStr + " and m.cancelyn = 'N'"
''        sqlStr = sqlStr + " and m.jumundiv <> '9'"
''        if (FRectDesignerID <>"") then
''			sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
''		end if
''        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''		sqlStr = sqlStr + " order by d.makerid asc"
''
''		rsget.PageSize = FPageSize
''
''		rsget.Open sqlStr,dbget,1
''		FTotalCount = rsget.RecordCount
''
''		if (FCurrPage * FPageSize < FTotalCount) then
''			FResultCount = FPageSize
''		else
''			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
''		end if
''
''
''		FPageCount = rsget.PageCount
''		FTotalPage = (FTotalCount\FPageSize)
''
''		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
''		redim preserve FMasterItemList(FResultCount)
''
''		if not rsget.EOF then
''			rsget.absolutepage = FCurrPage
''
''			do until (i >= FResultCount)
''
''				set FMasterItemList(i) = new CUpchebeasongDetail
''
''				FMasterItemList(i).FOrderserial = rsget("orderserial")
''				FMasterItemList(i).FItemid 	 = rsget("itemid")
''				FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
''				FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
''				FMasterItemList(i).FItemcnt     = rsget("itemno")
''				FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
''				FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
''				FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
''				FMasterItemList(i).FRegdate  = rsget("regdate")
''				FMasterItemList(i).FCurrstate  = rsget("baljuok")
''				FMasterItemList(i).FMakerid = rsget("makerid")
''
''				rsget.movenext
''				i=i+1
''
''			loop
''		end if
''		rsget.Close
''	end sub

''	public Sub DesignerDateMiBaljuCount()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select count(*) as cnt"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,"
''		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and m.regdate > '" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
''		sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
''		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''        sqlStr = sqlStr + " and m.cancelyn = 'N'"
''        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
''        sqlStr = sqlStr + " and m.jumundiv < 9"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''        sqlStr = sqlStr + " and ((d.currstate is null) or (d.currstate = 2))"
''
''		rsget.Open sqlStr,dbget,1
''
''		FResultCount = rsget.RecordCount
''		redim preserve FMasterItemList(0)
''
''		if Not rsget.Eof then
''			set FMasterItemList(0) = new CBaljuMaster
''			FMasterItemList(0).FTotalea = rsget("cnt")
''		end if
''
''		rsget.Close
''	end sub

''	public Sub DesignerDateMiBaljuList()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select distinct d.makerid, count(d.idx) as cnt"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,"
''		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
''		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''
''		if FRectDesignerID<>"" then
''			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
''		end if
''
''        sqlStr = sqlStr + " and m.cancelyn = 'N'"
''        sqlStr = sqlStr + " and m.jumundiv < 9"
''        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''        sqlStr = sqlStr + " and ((d.currstate is null) or (d.currstate = 2))"
''		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
''
''		sqlStr = sqlStr + " group by d.makerid"
''		sqlStr = sqlStr + " order by cnt desc"
''
''
'''response.write sqlStr
''
''		rsget.Open sqlStr,dbget,1
''
''		FResultCount = rsget.RecordCount
''		redim preserve FMasterItemList(FResultCount)
''			do until (i >= FResultCount)
''
''				set FMasterItemList(i) = new CBaljuMaster
''
''    			FMasterItemList(i).FMakerid = rsget("makerid")
''    			FMasterItemList(i).FTotalea = rsget("cnt")
''
''				rsget.movenext
''				i=i+1
''			loop
''		rsget.Close
''	end sub


''	public Sub UpchebeasongMibaljuList()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select d.itemno, m.orderserial, d.itemid, d.itemname,"
''		sqlStr = sqlStr + " d.itemoptionname, d.makerid, isNull(d.currstate,0) as baljuok,"
''		sqlStr = sqlStr + " m.cancelyn, m.buyname, m.reqname, m.regdate,m.ipkumdate"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m, " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
''        sqlStr = sqlStr + " and m.cancelyn = 'N'"
''        sqlStr = sqlStr + " and m.jumundiv < 9"
''        sqlStr = sqlStr + " and m.ipkumdiv >= 5"
''        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''        sqlStr = sqlStr + " and ((d.currstate is null) or (d.currstate = 2))"
''		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
''		sqlStr = sqlStr + " order by m.ipkumdate ,m.idx "
'''response.write sqlStr
''		rsget.PageSize = FPageSize
''
''		rsget.Open sqlStr,dbget,1
''		FTotalCount = rsget.RecordCount
''
''
''		if (FCurrPage * FPageSize < FTotalCount) then
''			FResultCount = FPageSize
''		else
''			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
''		end if
''
''
''		FPageCount = rsget.PageCount
''
''		FTotalPage = (FTotalCount\FPageSize)
''
''		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
''
''		redim preserve FMasterItemList(FResultCount)
''
''		if not rsget.EOF then
''			rsget.absolutepage = FCurrPage
''
''			do until (i >= FResultCount)
''
''				set FMasterItemList(i) = new CUpchebeasongDetail
''
''    			FMasterItemList(i).FOrderserial = rsget("orderserial")
''    			FMasterItemList(i).FItemid 	    = rsget("itemid")
''    			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
''    			FMasterItemList(i).FItemoption  = db2html(rsget("itemoptionname"))
''    			FMasterItemList(i).FItemcnt     = rsget("itemno")
''    			FMasterItemList(i).FBuyname     = db2html(rsget("buyname"))
''    			FMasterItemList(i).FReqname     = db2html(rsget("reqname"))
''    			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
''    			FMasterItemList(i).FRegdate    = rsget("regdate")
''    			FMasterItemList(i).FIpkumdate  = rsget("ipkumdate")
''    			FMasterItemList(i).Fmakerid    = rsget("makerid")
''    			FMasterItemList(i).FCurrstate  = rsget("baljuok")
''
''				rsget.movenext
''				i=i+1
''			loop
''		end if
''		rsget.Close
''	end sub

'		public Sub DesignerDateMiBaljuDetail()
'		dim sqlStr
'		dim i
'
'		sqlStr = "select d.itemno, m.orderserial, d.makerid, d.itemid, d.itemname,"
'		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
'		sqlStr = sqlStr + " m.cancelyn, m.buyname, m.reqname, m.regdate, m.ipkumdiv, m.ipkumdate, m.baljudate"
'		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m, " & TABLE_ORDERDETAIL & " d"
'		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'		sqlStr = sqlStr + " and m.regdate >= '" & dateAdd("m",-1,FRectRegStart) & "'"
'        sqlStr = sqlStr + " and m.baljudate >= '" & FRectRegStart & "'"
'		sqlStr = sqlStr + " and m.baljudate <= '" & FRectRegEnd & "'"
'
'		if FRectDesignerID <>"" then
'			sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
'		end if
'
'		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
'    	sqlStr = sqlStr + " and m.cancelyn = 'N'"
'    	sqlStr = sqlStr + " and m.jumundiv < 9"
'
'    	if FRectIpkumdiv <>"" then
'			sqlStr = sqlStr + " and m.ipkumdiv >= '" & FRectIpkumdiv & "'"
'		end if
'
'    	sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
'    	sqlStr = sqlStr + " and ((d.currstate is null) or (d.currstate = 2))"
'		sqlStr = sqlStr + " order by m.ipkumdate, d.idx"
'
'		rsget.PageSize = FPageSize
''response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'		FTotalCount = rsget.RecordCount
'
'
'		if (FCurrPage * FPageSize < FTotalCount) then
'			FResultCount = FPageSize
'		else
'			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
'		end if
'
'
'		FPageCount = rsget.PageCount
'
'		FTotalPage = (FTotalCount\FPageSize)
'
'		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
'
'		redim preserve FMasterItemList(FResultCount)
'
'		if not rsget.EOF then
'			rsget.absolutepage = FCurrPage
'
'			do until (i >= FResultCount)
'
'				set FMasterItemList(i) = new CBaljuMaster
'
'    			FMasterItemList(i).FOrderserial = rsget("orderserial")
'    			FMasterItemList(i).FMakerid     = rsget("makerid")
'    			FMasterItemList(i).FItemid 	    = rsget("itemid")
'    			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
'    			FMasterItemList(i).FItemoption  = db2html(rsget("itemoptionname"))
'    			FMasterItemList(i).FItemcnt     = rsget("itemno")
'    			FMasterItemList(i).FBuyname     = db2html(rsget("buyname"))
'    			FMasterItemList(i).FReqname     = db2html(rsget("reqname"))
'    			FMasterItemList(i).FCancelYn	= rsget("cancelyn")
'    			FMasterItemList(i).FRegdate     = rsget("regdate")
'    			FMasterItemList(i).Fipkumdate   = rsget("ipkumdate")
'    			FMasterItemList(i).FCurrstate   = rsget("baljuok")
'
'                FMasterItemList(i).Fbaljudate   = rsget("baljudate")
'
'
'				rsget.movenext
'				i=i+1
'			loop
'		end if
'		rsget.Close
'	end sub
'
'	public Sub DesignerDateMiBeasongCount()
'		dim sqlStr
'		dim i
'
'		sqlStr = "select count(*) as cnt"
'		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,"
'		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d"
'		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'		sqlStr = sqlStr + " and m.regdate >'" & FRectRegStart & "'"
'		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
'		sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
'		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
'		sqlStr = sqlStr + " and m.ipkumdiv >= 4"
'        sqlStr = sqlStr + " and m.jumundiv < 9"
'        sqlStr = sqlStr + " and m.cancelyn = 'N'"
'        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
'        sqlStr = sqlStr + " and d.currstate = '3'"
'
'		rsget.Open sqlStr,dbget,1
'
'		FResultCount = rsget.RecordCount
'		redim preserve FMasterItemList(0)
'
'		if Not rsget.Eof then
'			set FMasterItemList(0) = new CBaljuMaster
'			FMasterItemList(0).FTotalea = rsget("cnt")
'		end if
'
'		rsget.Close
'	end sub
'
'	public Sub DesignerDateMiBeasongNdayList()
'		dim sqlStr
'		dim i
'
'		sqlStr = "select d.makerid, "
'		sqlStr = sqlStr + " count(d.idx) as mibeasongcnt,"
'		sqlStr = sqlStr + " u.userdiv, u.socname_kor"
'		sqlStr = sqlStr + " from "
'		sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
'		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
'		sqlStr = sqlStr + " " & TABLE_USER_C & " u"
'		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
'		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
'		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=5"
'        sqlStr = sqlStr + " and m.cancelyn = 'N'"
'        sqlStr = sqlStr + " and m.jumundiv <> 9"
'        sqlStr = sqlStr + " and m.ipkumdiv > 3"
'        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
'        sqlStr = sqlStr + " and d.currstate <>'7'"
'        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
'        sqlStr = sqlStr + " and d.makerid=u.userid"
'        sqlStr = sqlStr + " and u.userdiv<14"
'
'		if FRectDesignerID<>"" then
'			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
'		end if
'
'		sqlStr = sqlStr + " group by d.makerid, u.userdiv, u.socname_kor"
'		sqlStr = sqlStr + " order by d.makerid Asc"
'
'		rsget.Open sqlStr,dbget,1
'
'		FResultCount = rsget.RecordCount
'		redim preserve FMasterItemList(FResultCount)
'			do until (i >= FResultCount)
'
'				set FMasterItemList(i) = new CUpCheSMSItem
'
'				FMasterItemList(i).FMakerid         = rsget("makerid")
'				FMasterItemList(i).FNDayMiBeasongCnt  = rsget("mibeasongcnt")
'
'				FMasterItemList(i).FUserDiv       = rsget("userdiv")
'				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))
'
'				rsget.movenext
'				i=i+1
'			loop
'		rsget.Close
'	end sub
'
'
'	public Sub DesignerDateMiBaljuNdayList()
'		dim sqlStr
'		dim i
'
'		sqlStr = "select d.makerid, "
'		sqlStr = sqlStr + " count(d.idx) as ndaymibaljucnt,"
'		sqlStr = sqlStr + " u.userdiv, u.socname_kor"
'		sqlStr = sqlStr + " from "
'		sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
'		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
'		sqlStr = sqlStr + " " & TABLE_USER_C & " u"
'		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
'		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
'		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=2"
'        sqlStr = sqlStr + " and m.cancelyn = 'N'"
'        sqlStr = sqlStr + " and m.jumundiv <> 9"
'        sqlStr = sqlStr + " and m.ipkumdiv > 3"
'        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
'        sqlStr = sqlStr + " and ((d.currstate is NULL) or (d.currstate = 2))"
'        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
'        sqlStr = sqlStr + " and d.makerid=u.userid"
'        sqlStr = sqlStr + " and u.userdiv<14"
'
'		if FRectDesignerID<>"" then
'			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
'		end if
'
'		sqlStr = sqlStr + " group by d.makerid, u.userdiv, u.socname_kor"
'		sqlStr = sqlStr + " order by d.makerid Asc"
'
'		rsget.Open sqlStr,dbget,1
'
'		FResultCount = rsget.RecordCount
'		redim preserve FMasterItemList(FResultCount)
'			do until (i >= FResultCount)
'
'				set FMasterItemList(i) = new CUpCheSMSItem
'
'				FMasterItemList(i).FMakerid         = rsget("makerid")
'				FMasterItemList(i).FNDayMiBaljuCnt = rsget("ndaymibaljucnt")
'
'				FMasterItemList(i).FUserDiv       = rsget("userdiv")
'				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))
'
'				rsget.movenext
'				i=i+1
'			loop
'		rsget.Close
'	end sub



''	public Sub DesignerDateMiBeasongList()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select d.makerid, count(d.idx) as cnt"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,"
''		sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
''		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''        sqlStr = sqlStr + " and m.cancelyn = 'N'"
''        sqlStr = sqlStr + " and m.jumundiv < 9"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''        sqlStr = sqlStr + " and d.currstate = '3'"
''		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
''
''		if FRectDesignerID<>"" then
''			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
''		end if
''
''		sqlStr = sqlStr + " group by d.makerid"
''		sqlStr = sqlStr + " order by cnt desc"
''
''		rsget.Open sqlStr,dbget,1
''
''		FResultCount = rsget.RecordCount
''		redim preserve FMasterItemList(FResultCount)
''			do until (i >= FResultCount)
''
''				set FMasterItemList(i) = new CBaljuMaster
''
''			FMasterItemList(i).FMakerid = rsget("makerid")
''			FMasterItemList(i).FTotalea = rsget("cnt")
''
''				rsget.movenext
''				i=i+1
''			loop
''		rsget.Close
''	end sub
''
''	public Sub DesignerDateMiBeasongDetailList()
''		dim sqlStr
''		dim i
''
''		sqlStr = "select d.itemno, m.orderserial, d.itemid, d.itemname,d.itemoptionname,"
''		sqlStr = sqlStr + " isNull(d.currstate,0) as baljuok,"
''		sqlStr = sqlStr + " m.cancelyn, m.regdate, m.buyname, m.reqname, m.ipkumdate, m.baljudate, d.upcheconfirmdate"
''		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m,  " & TABLE_ORDERDETAIL & " d"
''		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and m.regdate >= '" & dateAdd("m",-1,FRectRegStart) & "'"
''		sqlStr = sqlStr + " and m.ipkumdate >= '" & FRectRegStart & "'"
''		sqlStr = sqlStr + " and m.ipkumdate < '" & FRectRegEnd & "'"
''		sqlStr = sqlStr + " and m.cancelyn = 'N'"
''		sqlStr = sqlStr + " and m.jumundiv <> 9"
''		sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
''		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
''        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
''        sqlStr = sqlStr + " and d.currstate = '3'"
''		sqlStr = sqlStr + " order by m.baljudate desc, d.idx"
''
''		rsget.PageSize = FPageSize
''
''		rsget.Open sqlStr,dbget,1
''		FTotalCount = rsget.RecordCount
''
''
''		if (FCurrPage * FPageSize < FTotalCount) then
''			FResultCount = FPageSize
''		else
''			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
''		end if
''
''
''		FPageCount = rsget.PageCount
''
''		FTotalPage = (FTotalCount\FPageSize)
''
''		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
''
''		redim preserve FMasterItemList(FResultCount)
''
''		if not rsget.EOF then
''			rsget.absolutepage = FCurrPage
''
''			do until (i >= FResultCount)
''
''				set FMasterItemList(i) = new CBaljuMaster
''
''    			FMasterItemList(i).FOrderserial = rsget("orderserial")
''    			FMasterItemList(i).FItemid 	 = rsget("itemid")
''    			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
''    			FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
''    			FMasterItemList(i).FItemcnt     = rsget("itemno")
''    			FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
''    			FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
''    			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
''    			FMasterItemList(i).FRegdate  = rsget("regdate")
''    			FMasterItemList(i).Fipkumdate  = rsget("ipkumdate")
''    			FMasterItemList(i).FCurrstate  = rsget("baljuok")
''
''                FMasterItemList(i).Fbaljudate  = rsget("baljudate")
''                FMasterItemList(i).Fupcheconfirmdate = rsget("upcheconfirmdate")
''
''				rsget.movenext
''				i=i+1
''			loop
''		end if
''		rsget.Close
''	end sub

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

public function GetMichulgoSMSString(misendReason)
	select Case misendReason
		'// 출고지연
		CASE "03" : GetMichulgoSMSString = "[핑거스 출고지연안내]주문하신 상품 중 [상품명]([상품코드]) 상품이 [출고예정일]에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."

		'// 주문제작
		CASE "02" : GetMichulgoSMSString = "[핑거스 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 주문제작 상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 수입
		CASE "08" : GetMichulgoSMSString = "[핑거스 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 수입상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 가구배송
		CASE "09" : GetMichulgoSMSString = "[핑거스 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 가구상품으로 [출고예정일]에 발송될 예정이며, 우천시 변경될 수 있음 양해 부탁드립니다. 직배송 상품으로 배송 전 별도 연락드릴 예정입니다. 감사합니다."

		'// 예약배송
		CASE "04" : GetMichulgoSMSString = "[핑거스 출고예정안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 예약배송 상품으로 [출고예정일]에 발송될 예정입니다. 이와 관련, 문의사항은 고객센터로 연락 부탁드립니다. 감사합니다."

		'// 업체휴가
		CASE "10" : GetMichulgoSMSString = "[핑거스 출고지연안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 업체 휴가로 인해 [출고예정일]에 발송될 예정입니다. 빠른 배송 드리지 못해 죄송합니다."

		'// 고객지정배송
		CASE "07" : GetMichulgoSMSString = "[핑거스 출고 일정 안내]주문하신 상품 중 [상품명]([상품코드]) 상품은 고객지정배송상품으로 [출고예정일]에 발송될 예정입니다. 감사합니다."

		CASE ELSE : GetMichulgoSMSString = ""
	end Select
end function

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
		CASE "03" : GetMichulgoMailTitleString = "[핑거스] 출고지연안내 메일입니다."

		'// 주문제작
		CASE "02" : GetMichulgoMailTitleString = "[핑거스] 출고 일정 안내 메일입니다."

		'// 수입
		CASE "08" : GetMichulgoMailTitleString = "[핑거스] 출고 일정 안내 메일입니다."

		'// 가구배송
		CASE "09" : GetMichulgoMailTitleString = "[핑거스] 출고 일정 안내 메일입니다."

		'// 예약배송
		CASE "04" : GetMichulgoMailTitleString = "[핑거스] 출고예정안내 메일입니다."

		'// 업체휴가
		CASE "10" : GetMichulgoMailTitleString = "[핑거스] 출고지연안내 메일입니다."

		'// 고객지정배송
		CASE "07" : GetMichulgoMailTitleString = "[핑거스] 출고 일정 안내 메일입니다."

		CASE ELSE : GetMichulgoMailTitleString = ""
	end Select
end function








%>
