<%
Class CUpCheSMSItem
	public FMakerid
	public FCompanyName
	public FMiBalJuCount
	public FMiBeasongCount
	public FLastSendMsgDay
	public FDeliverHp
	public FUserDiv
	public FSocNameKor

	public FNDayMiBaljuCnt
	public FNDayMiBeasongCnt

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


	public FItemcnt
	public FJumunDiv

	public FBuyCash
	public FSellcash

	public FUpcheBeasongDate
	public Fmasteridx

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

	public FRectDateType
	public FRectDeliverType

    public FOrderserial
	public Fitemid
	public FItemname
	public FItemoptioncode
	public FItemoption
	public FItemcnt
	public FBuyname
	public FReqname
	public FCancelYn
	public FRegdate
	public FCurrstate
	public Fidx
    public Fipkumdiv
	public FMakerid
	public FTotalea

	public FIpkumDate
	public FRect



	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		'redim preserve FDetailItemList(0)

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

		sqlStr = sqlStr + " (select distinct top 2000 m.orderserial,"
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
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
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
			wheredetail = " and d.makerid='" & FRectDesignerID & "'"
		end if

		if (FRectDateType="ipkumil") then
			wheredetail = wheredetail + " and m.ipkumdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.ipkumdate < '" & FRectRegEnd & "'"
		else
			wheredetail = wheredetail + " and m.regdate >= '" & FRectRegStart & "'"
			wheredetail = wheredetail + " and m.regdate < '" & FRectRegEnd & "'"
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
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + wheredetail

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.close
		''#################################################
		''데이타.
		''#################################################

		sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " m.orderserial, m.jumundiv, d.itemno, d.itemid, d.itemname, d.buycash, d.itemcost,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.ipkumdate, m.regdate, m.buyname, m.reqname , d.makerid, d.cancelyn as detailcancelyn, d.beasongdate, d.isupchebeasong as deliverytype"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.itemid<>0"
        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " order by m.orderserial desc, d.makerid asc"
		rsget.PageSize = FPageSize

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1

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

	public Sub DesignerDateBaljuList()
		dim sqlStr
		dim i

		sqlStr = "select m.orderserial, d.itemno, d.itemid, d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.regdate, m.buyname, m.reqname , d.makerid"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,  [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
 		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv <> '9'"
        if (FRectDesignerID <>"") then
			sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		end if
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
		sqlStr = sqlStr + " order by d.makerid asc"

		rsget.PageSize = FPageSize

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
		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

				FMasterItemList(i).FOrderserial = rsget("orderserial")
				FMasterItemList(i).FItemid 	 = rsget("itemid")
				FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
				FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
				FMasterItemList(i).FItemcnt     = rsget("itemno")
				FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
				FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
				FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
				FMasterItemList(i).FRegdate  = rsget("regdate")
				FMasterItemList(i).FCurrstate  = rsget("baljuok")
				FMasterItemList(i).FMakerid = rsget("makerid")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close
	end sub

	public Sub DesignerDateMiBaljuCount()
		dim sqlStr
		dim i

        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibalju_Count '" + Cstr(FRectDesignerID) +  "'"

		rsget.Open sqlStr, dbget, 1
		FTotalea = rsget("cnt")
		rsget.Close
	end sub

	public Sub DesignerDateMiBaljuList()
		dim sqlStr
		dim i

		sqlStr = "select distinct d.makerid, count(d.idx) as cnt"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"

		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if

        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv < 9"
        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + " and d.currstate is null"
		sqlStr = sqlStr + " and d.oitemdiv<>'90'"

		sqlStr = sqlStr + " group by d.makerid"
		sqlStr = sqlStr + " order by cnt desc"


'response.write sqlStr

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

			FMasterItemList(i).FMakerid = rsget("makerid")
			FMasterItemList(i).FTotalea = rsget("cnt")

				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub


	public Sub UpchebeasongMibaljuList()
		dim sqlStr
		dim i

		sqlStr = "select d.itemno, m.orderserial, d.itemid, d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname, d.makerid, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.buyname, m.reqname, m.regdate,m.ipkumdate"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate < '" & FRectRegEnd & "'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv < 9"
        sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + " and d.currstate is null"
		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
		sqlStr = sqlStr + " order by m.ipkumdate ,m.idx "
'response.write sqlStr
		rsget.PageSize = FPageSize

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

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

			FMasterItemList(i).FOrderserial = rsget("orderserial")
			FMasterItemList(i).FItemid 	 = rsget("itemid")
			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
			FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
			FMasterItemList(i).FItemcnt     = rsget("itemno")
			FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
			FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
			FMasterItemList(i).FRegdate  = rsget("regdate")
			FMasterItemList(i).FIpkumdate  = rsget("ipkumdate")
			FMasterItemList(i).Fmakerid = rsget("makerid")
			FMasterItemList(i).FCurrstate  = rsget("baljuok")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

		public Sub DesignerDateMiBaljuDetail()
		dim sqlStr
		dim i


		''#################################################
		''데이타.
		''#################################################

		sqlStr = "select d.itemno, m.orderserial, d.itemid, d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.buyname, m.reqname, m.regdate, m.ipkumdate"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        	sqlStr = sqlStr + " and m.cancelyn = 'N'"
        	sqlStr = sqlStr + " and m.jumundiv < 9"
        	sqlStr = sqlStr + " and m.ipkumdiv >= 4"
        	sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        	sqlStr = sqlStr + " and d.currstate is null"
 		sqlStr = sqlStr + " and m.ipkumdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.ipkumdate < '" & FRectRegEnd & "'"
		sqlStr = sqlStr + " order by m.ipkumdate desc"
'response.write sqlStr
		rsget.PageSize = FPageSize

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

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

			FMasterItemList(i).FOrderserial = rsget("orderserial")
			FMasterItemList(i).FItemid 	 = rsget("itemid")
			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
			FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
			FMasterItemList(i).FItemcnt     = rsget("itemno")
			FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
			FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
			FMasterItemList(i).FRegdate  = rsget("regdate")
			FMasterItemList(i).Fipkumdate  = rsget("ipkumdate")
			FMasterItemList(i).FCurrstate  = rsget("baljuok")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub DesignerDateMiBeasongCount()
		dim sqlStr
		dim i
        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibeasong_Count '" + CStr(FRectDesignerID) +  "'"  + vbcrlf

		rsget.Open sqlStr, dbget, 1
		FTotalea = rsget("cnt")
		rsget.Close
	end sub

	public Sub DesignerDateMiBeasongNdayList()
		dim sqlStr
		dim i

		sqlStr = "select d.makerid, "
		sqlStr = sqlStr + " count(d.idx) as mibeasongcnt,"
		sqlStr = sqlStr + " u.userdiv, u.socname_kor"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c u"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=5"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv <> 9"
        sqlStr = sqlStr + " and m.ipkumdiv > 3"
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.currstate <>'7'"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + " and d.makerid=u.userid"
        sqlStr = sqlStr + " and u.userdiv<14"

		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if

		sqlStr = sqlStr + " group by d.makerid, u.userdiv, u.socname_kor"
		sqlStr = sqlStr + " order by d.makerid Asc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpCheSMSItem

				FMasterItemList(i).FMakerid         = rsget("makerid")
				FMasterItemList(i).FNDayMiBeasongCnt  = rsget("mibeasongcnt")

				FMasterItemList(i).FUserDiv       = rsget("userdiv")
				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))

				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub


	public Sub DesignerDateMiBaljuNdayList()
		dim sqlStr
		dim i

		sqlStr = "select d.makerid, "
		sqlStr = sqlStr + " count(d.idx) as ndaymibaljucnt,"
		sqlStr = sqlStr + " u.userdiv, u.socname_kor"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c u"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=2"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv <> 9"
        sqlStr = sqlStr + " and m.ipkumdiv > 3"
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.currstate is NULL"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + " and d.makerid=u.userid"
        sqlStr = sqlStr + " and u.userdiv<14"

		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if

		sqlStr = sqlStr + " group by d.makerid, u.userdiv, u.socname_kor"
		sqlStr = sqlStr + " order by d.makerid Asc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpCheSMSItem

				FMasterItemList(i).FMakerid         = rsget("makerid")
				FMasterItemList(i).FNDayMiBaljuCnt = rsget("ndaymibaljucnt")

				FMasterItemList(i).FUserDiv       = rsget("userdiv")
				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))

				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub


	public Sub DesignerDateMiBaljuMiBeasongList()
		dim sqlStr
		dim i

		sqlStr = "select d.makerid, "
		sqlStr = sqlStr + " sum(case when d.currstate is NULL then 1 else 0 end) as mibaljucnt,"
		sqlStr = sqlStr + " sum(case when (datediff(d,m.ipkumdate,getdate())>=2) and (d.currstate is NULL) then 1 else 0 end ) as ndaymibaljucnt,"
		sqlStr = sqlStr + " sum(case when (datediff(d,m.ipkumdate,getdate())>=5) and (d.currstate ='3') then 1 else 0 end ) as ndaymibeasongcnt,"
		sqlStr = sqlStr + " sum(case when d.currstate='3' then 1 else 0 end) as mibeasongcnt,"
		sqlStr = sqlStr + " u.userdiv, u.socname_kor, p.company_name, p.deliver_hp"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c u"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on u.userid=p.id"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
        sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.jumundiv<>9"
        sqlStr = sqlStr + " and m.ipkumdiv>3"
        sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " and d.makerid=u.userid"
        sqlStr = sqlStr + " and u.userdiv<14"
		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if
		'sqlStr = sqlStr + " and (sum(case when (datediff(d,m.ipkumdate,getdate())>=2) and (d.currstate is NULL) then 1 else 0 end )>0"
		'sqlStr = sqlStr + " 	or sum(case when (datediff(d,m.ipkumdate,getdate())>=2) and (d.currstate is NULL) then 1 else 0 end )>0)"

		sqlStr = sqlStr + " group by d.makerid, u.userdiv, u.socname_kor, p.company_name, p.deliver_hp"
		sqlStr = sqlStr + " order by d.makerid Asc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CUpCheSMSItem

				FMasterItemList(i).FMakerid         = rsget("makerid")

				FMasterItemList(i).FNDayMiBaljuCnt = rsget("ndaymibaljucnt")
				FMasterItemList(i).FMiBalJuCount    = rsget("mibaljucnt")
				FMasterItemList(i).FNDayMiBeasongCnt = rsget("ndaymibeasongcnt")
				FMasterItemList(i).FMiBeasongCount  = rsget("mibeasongcnt")

				FMasterItemList(i).FUserDiv       = rsget("userdiv")
				FMasterItemList(i).FSocNameKor       = db2html(rsget("socname_kor"))

				FMasterItemList(i).FCompanyName    = db2html(rsget("company_name"))
				FMasterItemList(i).FDeliverHp       = db2html(rsget("deliver_hp"))

				'FMasterItemList(i).FLastSendMsgDay  = rsget("")
				'if IsNULL(FMasterItemList(i).FDeliverHp) or (FMasterItemList(i).FDeliverHp="") then
				'	FMasterItemList(i).FDeliverHp       = rsget("manager_hp")
				'end if
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub





	public Sub DesignerDateMiBeasongList()
		dim sqlStr
		dim i

		sqlStr = "select d.makerid, count(d.idx) as cnt"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate >='" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.regdate <'" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        sqlStr = sqlStr + " and m.cancelyn = 'N'"
        sqlStr = sqlStr + " and m.jumundiv < 9"
        sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        sqlStr = sqlStr + " and d.currstate = '3'"
		sqlStr = sqlStr + " and d.oitemdiv<>'90'"

		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if

		sqlStr = sqlStr + " group by d.makerid"
		sqlStr = sqlStr + " order by cnt desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

			FMasterItemList(i).FMakerid = rsget("makerid")
			FMasterItemList(i).FTotalea = rsget("cnt")

				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public Sub DesignerDateMiBeasongDetailList()
		dim sqlStr
		dim i

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select d.itemno, m.orderserial, d.itemid, d.itemname,d.itemoptionname,"
		sqlStr = sqlStr + " isNull(d.currstate,0) as baljuok,"
		sqlStr = sqlStr + " m.cancelyn, m.regdate, m.buyname, m.reqname, m.ipkumdate"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,  [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdate >= '" & FRectRegStart & "'"
		sqlStr = sqlStr + " and m.ipkumdate < '" & FRectRegEnd & "'"
		sqlStr = sqlStr + " and m.cancelyn = 'N'"
		sqlStr = sqlStr + " and m.jumundiv <> 9"
		sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
        	sqlStr = sqlStr + " and d.cancelyn <> 'Y'"
        	sqlStr = sqlStr + " and d.currstate = '3'"
		sqlStr = sqlStr + " order by m.ipkumdate desc"

		rsget.PageSize = FPageSize

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

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CBaljuMaster

			FMasterItemList(i).FOrderserial = rsget("orderserial")
			FMasterItemList(i).FItemid 	 = rsget("itemid")
			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
			FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
			FMasterItemList(i).FItemcnt     = rsget("itemno")
			FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
			FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
			FMasterItemList(i).FRegdate  = rsget("regdate")
			FMasterItemList(i).Fipkumdate  = rsget("ipkumdate")
			FMasterItemList(i).FCurrstate  = rsget("baljuok")

				rsget.movenext
				i=i+1
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
