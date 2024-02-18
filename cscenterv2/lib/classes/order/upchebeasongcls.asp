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
			GetMallName = "������"
		elseif FUserDiv="03" then
			GetMallName = "�ö��"
		elseif FUserDiv="04" then
			GetMallName = "�м�"
		elseif FUserDiv="05" then
			GetMallName = "���"
		elseif FUserDiv="06" then
			GetMallName = "��Ƽ"
		elseif FUserDiv="07" then
			GetMallName = "�ְ�"
		elseif FUserDiv="08" then
			GetMallName = "�������"
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
            CASE 0 : getMisendStateText="��ó��"
            CASE 4 : getMisendStateText="���ȳ�"
            CASE 6 : getMisendStateText="CSó���Ϸ�"
            CASE ELSE : getMisendStateText = FMisendState
        end Select
    end function

    public function getMisendText()
        select Case FMisendReason
            CASE "00" : getMisendText = "�Է´��"
            CASE "01" : getMisendText = "������"
            CASE "04" : getMisendText = "�����ǰ"

            CASE "02" : getMisendText = "�ֹ�����"
            CASE "52" : getMisendText = "�ֹ�����"
            CASE "03" : getMisendText = "�������"
            CASE "53" : getMisendText = "�������"
            CASE "05" : getMisendText = "ǰ�����Ұ�"
            CASE "55" : getMisendText = "ǰ�����Ұ�"
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

    public FRectCurrState	' ����

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
		''�Ѱ���
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
		''����Ÿ.
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


        ''[CS]��ü��۰���>>��ü��۰���
	public Sub DesignerDateMiBaljuMiBeasongList()
		dim sqlStr
		dim i
        ''���� ��ǰ�� ������(��������)�� ����

        sqlStr = " select T.*"
        sqlStr = sqlStr + " ,u.userdiv, u.socname_kor, u.catecode, cl.code_nm, p.company_name, p.deliver_hp, p.deliver_phone"
        sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + "     select d.makerid "
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '0') then 1 else 0 end) as mitongbocnt" ''���뺸 ��ü
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '2') then 1 else 0 end) as mibaljucnt"  ''��Ȯ�� ��ü
		sqlStr = sqlStr + "     ,sum(case when (d.currstate = '3') then 1 else 0 end) as mibeasongcnt" ''����� ��ü
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
        sqlStr = sqlStr + "     and m.ipkumdiv<'8'"                       ''���Ϸ� ����
        sqlStr = sqlStr + "     and d.itemid<>0"
        if FRectDesignerID<>"" then
			sqlStr = sqlStr + "     and d.makerid='" + CStr(FRectDesignerID) + "'"
		end if
        sqlStr = sqlStr + "     and d.isupchebeasong='Y'"
        sqlStr = sqlStr + "     and d.cancelyn<>'Y'"
		sqlStr = sqlStr + "     and d.currstate<'7'"                      ''���Ϸ� ���� / �ε��� ��Ÿ��.

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

	''[CS]��۰���>>�������Ʈ_���� ��
	public Sub getUpcheMichulgoList(byval isALL)
	    dim sqlStr
		dim i

		Dim stOrderSerial, edOrderserial
		stOrderSerial = Mid(Replace(CStr(FRectRegStart),"-",""),3,6) + "00000"
		edOrderserial = Mid(Replace(CStr(FRectRegEnd),"-",""),3,6) + "00000"

        '' baljudate => ��ǰ (�ֹ��뺸��=������) �� ����
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

	''[CS]��ü��۰���>>��ü��۸�� /�˾�
    public Sub getUpchebeasongList()
        dim sqlStr
		dim i
        '' baljudate => ��ǰ (�ֹ��뺸��=������) �� ����
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
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuCount()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuList()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub UpchebeasongMibaljuList()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuDetail()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongCount()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongNdayList()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBaljuNdayList()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongList()
'        response.write "������� - ������ ���� ���"
'        dbget.close()	:	response.End
'    end Sub
'
'    public Sub DesignerDateMiBeasongDetailList()
'        response.write "������� - ������ ���� ���"
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
		'// �������
		CASE "03" : GetMichulgoSMSString = "[�ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� [�������]�� �߼۵� �����Դϴ�. ���ο� ������ ��� �˼��մϴ�."

		'// �ֹ�����
		CASE "02" : GetMichulgoSMSString = "[�ΰŽ� ��� ���� �ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� �ֹ����� ��ǰ���� [�������]�� �߼۵� �����Դϴ�. �̿� ����, ���ǻ����� �����ͷ� ���� ��Ź�帳�ϴ�. �����մϴ�."

		'// ����
		CASE "08" : GetMichulgoSMSString = "[�ΰŽ� ��� ���� �ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� ���Ի�ǰ���� [�������]�� �߼۵� �����Դϴ�. �̿� ����, ���ǻ����� �����ͷ� ���� ��Ź�帳�ϴ�. �����մϴ�."

		'// �������
		CASE "09" : GetMichulgoSMSString = "[�ΰŽ� ��� ���� �ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� ������ǰ���� [�������]�� �߼۵� �����̸�, ��õ�� ����� �� ���� ���� ��Ź�帳�ϴ�. ����� ��ǰ���� ��� �� ���� �����帱 �����Դϴ�. �����մϴ�."

		'// ������
		CASE "04" : GetMichulgoSMSString = "[�ΰŽ� ������ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� ������ ��ǰ���� [�������]�� �߼۵� �����Դϴ�. �̿� ����, ���ǻ����� �����ͷ� ���� ��Ź�帳�ϴ�. �����մϴ�."

		'// ��ü�ް�
		CASE "10" : GetMichulgoSMSString = "[�ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� ��ü �ް��� ���� [�������]�� �߼۵� �����Դϴ�. ���� ��� �帮�� ���� �˼��մϴ�."

		'// ���������
		CASE "07" : GetMichulgoSMSString = "[�ΰŽ� ��� ���� �ȳ�]�ֹ��Ͻ� ��ǰ �� [��ǰ��]([��ǰ�ڵ�]) ��ǰ�� ��������ۻ�ǰ���� [�������]�� �߼۵� �����Դϴ�. �����մϴ�."

		CASE ELSE : GetMichulgoSMSString = ""
	end Select
end function

public function GetMichulgoMailString(misendReason)
	dim mailText

	mailText = ""
	select Case misendReason
		'// �������
		CASE "03" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.\n"
			mailText = mailText + "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.\n"
			mailText = mailText + "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,\n"
			mailText = mailText + "���ູ���ͷ� ���� ��Ź �帳�ϴ�.\n"
			mailText = mailText + "���ο� ������ �帰 �� �������� ��� �帮��, ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.\n"

		'// �ֹ�����
		CASE "02" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� �ֹ� �� ���� �Ǵ� ��ǰ����\n"
			mailText = mailText + "�Ϲݻ�ǰ�� �޸� �ֹ����ۿ� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.\n"
			mailText = mailText + "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,\n"
			mailText = mailText + "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.\n"

		'// ����
		CASE "08" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� ��ǰ ���� �� �߼۵Ǵ� ��ǰ����\n"
			mailText = mailText + "�Ϲݻ�ǰ�� �޸� ��ǰ ���Կ� ���� �� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.\n"
			mailText = mailText + "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,\n"
			mailText = mailText + "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.\n"

		'// �������
		CASE "09" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� ���� ��ǰ����\n"
			mailText = mailText + "�Ϲݻ�ǰ�� �޸� ��ۿ� ���� �� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.\n"
			mailText = mailText + "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,\n"
			mailText = mailText + "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.\n"
			mailText = mailText + "���� ��ۻ�ǰ���� ��õ�� ������ ���� ����� �� ������, \n"
			mailText = mailText + "����� ��ǰ���� ��� �� ���� ���� �帮�ڽ��ϴ�.\n"
			mailText = mailText + "�̿� ����, ���ǻ����� ���ູ���ͷ� ���� ��Ź�帳�ϴ�.\n"

		'// ������
		CASE "04" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.\n"
			mailText = mailText + "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,\n"
			mailText = mailText + "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,\n"
			mailText = mailText + "���ູ���ͷ� ���� ��Ź�帳�ϴ�.\n"

		'// ��ü�ް�
		CASE "10" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� ��ü �ް� �Ⱓ���� ���� �߼��� ������ �����Դϴ�.\n"
			mailText = mailText + "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.\n"
			mailText = mailText + "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,\n"
			mailText = mailText + "���ູ���ͷ� ���� ��Ź �帳�ϴ�.\n"
			mailText = mailText + "���ο� ������ �帰 �� �������� ��� �帮��, ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.\n"

		'// ���������
		CASE "07" :
			mailText = mailText + "�ȳ��ϼ���. ����\n\n"
			mailText = mailText + "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.\n"
			mailText = mailText + "�ֹ��Ͻ� ��ǰ�� ��������ۻ�ǰ���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,\n"
			mailText = mailText + "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,\n"
			mailText = mailText + "���ູ���ͷ� ���� ��Ź�帳�ϴ�.\n"

		CASE ELSE :
			mailText = ""

	end Select

	GetMichulgoMailString = mailText
end function

public function GetMichulgoMailTitleString(misendReason)
	select Case misendReason
		'// �������
		CASE "03" : GetMichulgoMailTitleString = "[�ΰŽ�] ��������ȳ� �����Դϴ�."

		'// �ֹ�����
		CASE "02" : GetMichulgoMailTitleString = "[�ΰŽ�] ��� ���� �ȳ� �����Դϴ�."

		'// ����
		CASE "08" : GetMichulgoMailTitleString = "[�ΰŽ�] ��� ���� �ȳ� �����Դϴ�."

		'// �������
		CASE "09" : GetMichulgoMailTitleString = "[�ΰŽ�] ��� ���� �ȳ� �����Դϴ�."

		'// ������
		CASE "04" : GetMichulgoMailTitleString = "[�ΰŽ�] ������ȳ� �����Դϴ�."

		'// ��ü�ް�
		CASE "10" : GetMichulgoMailTitleString = "[�ΰŽ�] ��������ȳ� �����Դϴ�."

		'// ���������
		CASE "07" : GetMichulgoMailTitleString = "[�ΰŽ�] ��� ���� �ȳ� �����Դϴ�."

		CASE ELSE : GetMichulgoMailTitleString = ""
	end Select
end function








%>
