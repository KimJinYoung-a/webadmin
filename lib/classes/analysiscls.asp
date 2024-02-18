
<%


Class COnlineDailyGainItem
    public Fyyyymmdd
    public Forgitemcost
    public Ftotalsum
    public Fmiletotalprice
    public Ftencardspend
    public Fallatdiscountprice
    public Fspendmembership
    public Fsubtotalprice
    public Fitemtotalsum
    public Fitembuysum
    public Fdeliverytotalsum

    public FtenbeaCount

    public function GetDeliveryBuyCash()
        if Fyyyymmdd>="2007-04-01" then
            GetDeliveryBuyCash = 1600
        elseif Fyyyymmdd>="2019-01-01" then
            GetDeliveryBuyCash = 2000
        else
            GetDeliveryBuyCash = 2500
        end if
    end function

    public function GetErrSubTotal()
        GetErrSubTotal = Ftotalsum-Fmiletotalprice-Ftencardspend-Fallatdiscountprice-Fspendmembership-Fsubtotalprice
    end function

    public function GetGainSum()
        ''일별수익 = 총매출 - 매입 - 텐배송건수 * 배송단가
        GetGainSum = Fsubtotalprice - Fitembuysum - FtenbeaCount*GetDeliveryBuyCash
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COnlineMonthGainItem
	public FWebTotalSel
    public FMeTotal
	public FWiTotal
	public FUbTotal
	public FEtTotal
    public FDlvTotal

	public function getTotalMeaip()
		getTotalMeaip = FMeTotal + FWiTotal + FUbTotal + FEtTotal + FDlvTotal
	end function

	Private Sub Class_Initialize()
		FWebTotalSel = 0
		FMeTotal    = 0
		FWiTotal    = 0
		FUbTotal    = 0
		FEtTotal    = 0
        FDlvTotal   = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class COffShopMonthGainItem
	public FShopID
	public Fchargediv
	public FChargeDivName
	public FSuplysum
	public Ftotsum
	public Fminuscharge
	public Frealjungsansum
	public FTotSpendMile

	public Fupchebuysum
	public Fshopsuplysum

	public Fchul_upchebuysum
	public Fchul_shopsuplysum
	public Fre_upchebuysum
	public Fre_shopsuplysum


	public function getJungSanChargeDivName()
		if Fchargediv="2" then
			getJungSanChargeDivName = "10x10 특정"
		elseif Fchargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif Fchargediv="5" then
			getJungSanChargeDivName = "출고 정산"
		elseif Fchargediv="6" then
			getJungSanChargeDivName = "업체 특정"
		elseif Fchargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		'elseif Fchargediv="T" then
		'	getJungSanChargeDivName = "매입->매입"
		else
			getJungSanChargeDivName = Fchargediv
		end if
	end function

	Private Sub Class_Initialize()
		Ftotsum = 0
		Fminuscharge = 0
		Frealjungsansum = 0
		FTotSpendMile = 0

		Fupchebuysum = 0
		Fshopsuplysum = 0

		Fchul_upchebuysum 	= 0
		Fchul_shopsuplysum   = 0
		Fre_upchebuysum 	= 0
		Fre_shopsuplysum 	= 0

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COnlineMonthlyGainItem
	public Fyyyymm
	public Fmcnt
	public FTotalSum
	public FSubTotalPrice

	public Fminuscnt
	public FminusTotalSum
	public FminusSubTotalPrice

	public Fmiletotalprice
	public Ftencardspend
	public Fspendmembership
	public Fallatdiscountprice

	public FMeaipTotal
	public FBeasongCnt
	public FBeasongPay

	public function getRealSubTotalPrice()
		getRealSubTotalPrice = FSubTotalPrice + FminusSubTotalPrice
	end function

	public function getTotalDiscountsum()
		getTotalDiscountsum = Fmiletotalprice + Ftencardspend + Fspendmembership + Fallatdiscountprice
	end function

	public Function GetSuic()
		GetSuic = FSubTotalPrice + FminusSubTotalPrice - FMeaipTotal - GetBeasongTotal
	end function

	public function GetBeasongTotal()
		GetBeasongTotal = FBeasongCnt * FBeasongPay
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CAnalysis
	public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
	public FRectYYYYMM
	public FRectYYYYMM2
	public FRectYYYYMMDD
	public FRectYYYYMMDD2

	public FBeasongPay
    public FRectOldList

	public function getIpjumSusu()
		if FRectShopID="streetshop002" then
			getIpjumSusu = 20
		elseif FRectShopID="streetshop099" then
			getIpjumSusu = 20
		else
			getIpjumSusu = 0
		end if

	end function

	public function getIpjumSusuSum()
		getIpjumSusuSum = CLng(GetTotalMeachul*GetIpjumSusu/100)

	end function

	public function GetTotalMeachul()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Ftotsum
			end if
		next
		GetTotalMeachul = re
	end function

	public function GetTotalMinusCharge()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Fminuscharge
			end if
		next
		GetTotalMinusCharge = re
	end function

	public function GetTotalRealSum()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Frealjungsansum
			end if
		next
		GetTotalRealSum = re
	end function

	public function GetTotalShopSuplySum()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Fshopsuplysum
			end if
		next
		GetTotalShopSuplySum = re
	end function

	public function GetTotalShop_ChulSuplySum()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Fchul_shopsuplysum
			end if
		next
		GetTotalShop_ChulSuplySum = re
	end function

	public function GetTotalShop_ReSuplySum()
		dim i,re
		re = 0
		for i=LBound(FItemList) to UBound(FItemList)-1
			if Not (FItemList(i) is Nothing) then
				re = re + FItemList(i).Fre_shopsuplysum
			end if
		next
		GetTotalShop_ReSuplySum = re
	end function

    public Sub getOnLineDailyGainSum()
        dim sqlStr,i
        sqlStr = "select T1.*, T2.*, T3.* " + VbCrlf
        sqlStr = sqlStr + " from " + VbCrlf
        sqlStr = sqlStr + " 	( " + VbCrlf
        sqlStr = sqlStr + " 	select convert(varchar(10),m.regdate,21)  as yyyymmdd, sum(m.totalsum) as totalsum, " + VbCrlf
        sqlStr = sqlStr + " 	sum(m.miletotalprice) as miletotalprice, " + VbCrlf
        sqlStr = sqlStr + " 	sum(m.tencardspend) as tencardspend, " + VbCrlf
        sqlStr = sqlStr + " 	sum(allatdiscountprice) as allatdiscountprice, " + VbCrlf
        sqlStr = sqlStr + " 	sum(spendmembership) as spendmembership," + VbCrlf
        sqlStr = sqlStr + " 	sum(m.subtotalprice) as subtotalprice" + VbCrlf
        if (FRectOldList="on") then
            sqlStr = sqlStr + " 	from [db_log].[dbo].tbl_old_order_master_2003 m" + VbCrlf
        else
            sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m" + VbCrlf
        end if
        sqlStr = sqlStr + " 	where m.regdate>='" + FRectYYYYMMDD + "'" + VbCrlf
        sqlStr = sqlStr + " 	and m.regdate<'" + FRectYYYYMMDD2 + "'" + VbCrlf
        sqlStr = sqlStr + " 	and m.ipkumdiv>3" + VbCrlf
        sqlStr = sqlStr + " 	and m.jumundiv<>9" + VbCrlf
        sqlStr = sqlStr + " 	and m.cancelyn='N'" + VbCrlf
        sqlStr = sqlStr + " 	group by convert(varchar(10),m.regdate,21) " + VbCrlf
        sqlStr = sqlStr + " 	) T1" + VbCrlf
        sqlStr = sqlStr + " 	left join (" + VbCrlf
        sqlStr = sqlStr + " 		select convert(varchar(10),m.regdate,21)  as yyyymmdd," + VbCrlf
        sqlStr = sqlStr + " 		sum(case when d.itemid<>0 then (d.orgitemcost*d.itemno) else 0 end ) as orgitemcost," + VbCrlf
        sqlStr = sqlStr + " 		sum(case when d.itemid<>0 then (d.itemcost*d.itemno) else 0 end ) as itemtotalsum," + VbCrlf
        sqlStr = sqlStr + " 		sum(case when d.itemid<>0 then (d.buycash*d.itemno) else 0 end ) as itembuysum," + VbCrlf
        sqlStr = sqlStr + " 		sum(case when d.itemid=0 then (d.itemcost*d.itemno) else 0 end ) as deliverytotalsum" + VbCrlf

        if (FRectOldList="on") then
            sqlStr = sqlStr + " 		from [db_log].[dbo].tbl_old_order_master_2003 m," + VbCrlf
            sqlStr = sqlStr + " 		 [db_log].[dbo].tbl_old_order_detail_2003 d" + VbCrlf
        else
            sqlStr = sqlStr + " 		from [db_order].[dbo].tbl_order_master m," + VbCrlf
            sqlStr = sqlStr + " 		 [db_order].[dbo].tbl_order_detail d" + VbCrlf
        end if

        sqlStr = sqlStr + " 		where m.orderserial=d.orderserial" + VbCrlf
        sqlStr = sqlStr + " 		and m.regdate>='" + FRectYYYYMMDD + "'" + VbCrlf
        sqlStr = sqlStr + " 		and m.regdate<'" + FRectYYYYMMDD2 + "'" + VbCrlf
        sqlStr = sqlStr + " 		and m.ipkumdiv>3" + VbCrlf
        sqlStr = sqlStr + " 		and m.jumundiv<>9" + VbCrlf
        sqlStr = sqlStr + " 		and m.cancelyn='N'" + VbCrlf
        sqlStr = sqlStr + " 		and d.cancelyn<>'Y'" + VbCrlf
        sqlStr = sqlStr + " 		group by convert(varchar(10),m.regdate,21) " + VbCrlf
        sqlStr = sqlStr + " 	) T2 on T1.yyyymmdd=T2.yyyymmdd" + VbCrlf
        sqlStr = sqlStr + " 	Left Join (" + VbCrlf
        sqlStr = sqlStr + " 		select convert(varchar(10),m.regdate,21)  as yyyymmdd," + VbCrlf
        sqlStr = sqlStr + " 		count(distinct d.orderserial) as tenbeaCount" + VbCrlf

        if (FRectOldList="on") then
            sqlStr = sqlStr + " 		from [db_log].[dbo].tbl_old_order_master_2003 m," + VbCrlf
            sqlStr = sqlStr + " 		 [db_log].[dbo].tbl_old_order_detail_2003 d" + VbCrlf
        else
            sqlStr = sqlStr + " 		from [db_order].[dbo].tbl_order_master m," + VbCrlf
            sqlStr = sqlStr + " 		 [db_order].[dbo].tbl_order_detail d" + VbCrlf
        end if
        sqlStr = sqlStr + " 		where m.orderserial=d.orderserial" + VbCrlf
        sqlStr = sqlStr + " 		and m.regdate>='" + FRectYYYYMMDD + "'" + VbCrlf
        sqlStr = sqlStr + " 		and m.regdate<'" + FRectYYYYMMDD2 + "'" + VbCrlf
        sqlStr = sqlStr + " 		and m.ipkumdiv>3" + VbCrlf
        sqlStr = sqlStr + " 		and m.jumundiv<>9" + VbCrlf
        sqlStr = sqlStr + " 		and m.cancelyn='N'" + VbCrlf
        sqlStr = sqlStr + " 		and d.cancelyn<>'Y'" + VbCrlf
        sqlStr = sqlStr + " 		and d.itemid<>0" + VbCrlf
        sqlStr = sqlStr + " 		and d.isupchebeasong='N'" + VbCrlf
        sqlStr = sqlStr + " 		group by convert(varchar(10),m.regdate,21) " + VbCrlf
        sqlStr = sqlStr + " 	) T3 on T1.yyyymmdd=T3.yyyymmdd" + VbCrlf
        sqlStr = sqlStr + " order by T1.yyyymmdd" + VbCrlf

'response.write sqlStr

        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COnlineDailyGainItem
			FItemList(i).Fyyyymmdd           = rsget("yyyymmdd")
			FItemList(i).Forgitemcost        = rsget("orgitemcost")
            FItemList(i).Ftotalsum           = rsget("totalsum")
            FItemList(i).Fmiletotalprice     = rsget("miletotalprice")
            FItemList(i).Ftencardspend       = rsget("tencardspend")
            FItemList(i).Fallatdiscountprice = rsget("allatdiscountprice")
            FItemList(i).Fspendmembership    = rsget("spendmembership")
            FItemList(i).Fsubtotalprice      = rsget("subtotalprice")
            FItemList(i).Fitemtotalsum       = rsget("itemtotalsum")
            FItemList(i).Fitembuysum         = rsget("itembuysum")
            FItemList(i).Fdeliverytotalsum   = rsget("deliverytotalsum")


			i=i+1
			rsget.moveNext
		loop

		rsget.Close


    end Sub

	public Sub GetMeachulWithCopons()
		dim sqlStr,i
		sqlStr = "select "
		sqlStr = sqlStr + " count(m.idx) as mcnt, sum(m.totalsum) as  totalsum, sum(m.subtotalprice) as subtotalprice,"
		sqlStr = sqlStr + " sum(m.miletotalprice) as miletotalprice, sum(IsNull(m.tencardspend,0)) as tencardspend,"
		sqlStr = sqlStr + " sum(spendmembership) as spendmembership,"
		sqlStr = sqlStr + " sum(allatdiscountprice) as allatdiscountprice"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and m.cancelyn='N'"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		set FOneItem = new COnlineMonthlyGainItem
		FOneItem.Fyyyymm		= Left(FRectYYYYMMDD,7)

		if Not rsget.Eof then
			FOneItem.Fmcnt		= rsget("mcnt")
			FOneItem.FTotalSum	= rsget("totalsum")
			FOneItem.FSubTotalPrice	= rsget("subtotalprice")
			FOneItem.Fmiletotalprice = rsget("miletotalprice")
			FOneItem.Ftencardspend	 = rsget("tencardspend")
			FOneItem.Fspendmembership	= rsget("spendmembership")
			FOneItem.Fallatdiscountprice = rsget("allatdiscountprice")
		end if

		rsget.Close

	end sub

	public Sub GetMinusMeachulSum()
		dim sqlStr,i
		sqlStr = "select "
		sqlStr = sqlStr + " count(m.idx) as minuscnt, sum(m.totalsum) as  minustotalsum, sum(m.subtotalprice) as minussubtotalprice"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.jumundiv=9"
		sqlStr = sqlStr + " and m.cancelyn='N'"

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			FOneItem.Fminuscnt		= rsget("minuscnt")
			FOneItem.FminusTotalSum	= rsget("minustotalsum")
			FOneItem.FminusSubTotalPrice	= rsget("minussubtotalprice")
		end if

		rsget.Close

	end sub

	public Sub GetTenBeasongcount()
		dim sqlStr,i

		sqlStr = "select count(T.orderserial) as beasongcnt from ("
		sqlStr = sqlStr + " 	select m.orderserial, count(d.idx) as totcnt, sum( case when d.isupchebeasong='Y' then 1 else 0 end) as upchecnt"
		sqlStr = sqlStr + "  	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " 	and m.regdate<'" + FRectYYYYMMDD2 + "'"
		sqlStr = sqlStr + " 	and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.ipkumdiv>3"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T	"
		sqlStr = sqlStr + " where T.totcnt>T.upchecnt"

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			FOneItem.FBeasongPay = FBeasongPay
			FOneItem.FBeasongCnt = rsget("beasongcnt")
		end if
		rsget.Close
	end Sub

	public Sub GetMeaipSum()
		''상품 매입가
		dim sqlStr,i
		sqlStr = "select "
		sqlStr = sqlStr + " sum(d.buycash*d.itemno) as meaipttl"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			FOneItem.FMeaipTotal = rsget("meaipttl")
		end if
		rsget.Close
	end sub

	public Sub getOnlineMonthlyGain()
		dim sqlStr,i
		i=0

		sqlStr = "select "
		sqlStr = sqlStr + " count(m.idx) as mcnt, sum(m.totalsum) as  totalsum, sum(m.subtotalprice) as subtotalprice,"
		sqlStr = sqlStr + " sum(m.miletotalprice) as miletotalprice, sum(IsNull(m.tencardspend,0)) as tencardspend,"
		sqlStr = sqlStr + " sum(case when jcnt>0 then 1 else 0 end) as beasongcnt"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join ("
			sqlStr = sqlStr + " select m.orderserial,"
			sqlStr = sqlStr + " Sum(case when (d.isupchebeasong='Y') or (d.beasongdate is Not NULL) then 0 else 1 end ) as jcnt"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
			sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
			sqlStr = sqlStr + " and m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " group by m.orderserial"
			sqlStr = sqlStr + " ) as T on m.orderserial=T.orderserial"

		sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
		'sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COnlineMonthlyGainItem
			FItemList(i).Fyyyymm		= Left(FRectYYYYMMDD,7)
			FItemList(i).Fmcnt		= rsget("mcnt")
			FItemList(i).FTotalSum	= rsget("totalsum")
			FItemList(i).FSubTotalPrice	= rsget("subtotalprice")
			FItemList(i).Fmiletotalprice = rsget("miletotalprice")
			FItemList(i).Ftencardspend	 = rsget("tencardspend")
			FItemList(i).FBeasongPay	= FBeasongPay
			FItemList(i).FBeasongCnt = rsget("beasongcnt")
			i=i+1
			rsget.moveNext
		loop

		rsget.Close



		sqlStr = "select convert(varchar(7),m.regdate,20) as yyyymm, "
		sqlStr = sqlStr + " sum(d.buycash*d.itemno) as meaipttl"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.regdate>='" + FRectYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectYYYYMMDD2 + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.sitename<>'tingmart'"
		'sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20)"
		sqlStr = sqlStr + " order by yyyymm"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof
			for i=0 to FResultCount-1
				if (FItemList(i).Fyyyymm=rsget("yyyymm")) then
					FItemList(i).FMeaipTotal = rsget("meaipttl")

				end if
			next
			rsget.moveNext
		loop

		rsget.Close

	end Sub

	public Sub getOnLineMonthGainSum()
	    dim isOldLogData
	    dim tmpStartDate, tmpEndDate
	    tmpStartDate = FRectYYYYMM + "-01"
	    tmpEndDate   = CStr(DateAdd("m",1,tmpStartDate))

	    isOldLogData = datediff("m",FRectYYYYMM & "-01",now())>6

		dim sqlStr,i
		sqlStr = "select sum(subtotalprice) as subtotalprice from "
		if isOldLogData then
		    sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m"
		else
		    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
	    end if

		sqlStr = sqlStr + " where regdate>='" + tmpStartDate + "'"
		sqlStr = sqlStr + " and regdate<'" + tmpEndDate + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and ipkumdiv>3"

		rsget.Open sqlStr,dbget,1

		set FOneItem = new COnlineMonthGainItem
		if Not rsget.Eof then
			FOneItem.FWebTotalSel = rsget("subtotalprice")
			if IsNULL(FOneItem.FWebTotalSel) then FOneItem.FWebTotalSel = 0
		else
			FOneItem.FWebTotalSel = 0
		end if
		rsget.Close

		sqlStr = "select sum(me_totalsuplycash) as mettl,"
		sqlStr = sqlStr + " sum(wi_totalsuplycash) as wittl,"
		sqlStr = sqlStr + " sum(ub_totalsuplycash) as ubttl,"
		sqlStr = sqlStr + " sum(et_totalsuplycash) as etttl"
		sqlStr = sqlStr + " sum(dlv_totalsuplycash) as dlvttl"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and cancelyn='N'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			FOneItem.FMeTotal = rsget("mettl")
			FOneItem.FWiTotal = rsget("wittl")
			FOneItem.FUbTotal = rsget("ubttl")
			FOneItem.FEtTotal = rsget("etttl")
			FOneItem.FDlvTotal = rsget("dlvttl")

			if IsNULL(FOneItem.FMeTotal) then FOneItem.FMeTotal = 0
			if IsNULL(FOneItem.FWiTotal) then FOneItem.FWiTotal = 0
			if IsNULL(FOneItem.FUbTotal) then FOneItem.FUbTotal = 0
			if IsNULL(FOneItem.FEtTotal) then FOneItem.FEtTotal = 0
			if IsNULL(FOneItem.FDlvTotal) then FOneItem.FDlvTotal = 0
		end if
		rsget.Close


		''오프샾 매입(10x10 매입)
'		sqlStr = " select m.socid, IsNull(sum(d.sellcash*d.itemno*-1),0) as ttlsellcash,"
'		sqlStr = sqlStr + " IsNull(sum(d.buycash*d.itemno*-1),0) as ttlbuycash"
'		sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m"
'		sqlStr = sqlStr + " , [db_storage].[dbo].tbl_acount_storage_detail d"
'		sqlStr = sqlStr + "  , [db_item].[dbo].tbl_item i"
'		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_user s"
'		sqlStr = sqlStr + " where convert(varchar(7),m.executedt,20)='" + FRectYYYYMM + "'"
'		sqlStr = sqlStr + " and m.code=d.mastercode"
'		sqlStr = sqlStr + " and m.deldt is NULL"
'		sqlStr = sqlStr + " and m.socid=s.userid"
'		sqlStr = sqlStr + " and d.deldt is NULL"
'		sqlStr = sqlStr + " and d.itemid=i.itemid and d.iitemgubun='10'"
'		sqlStr = sqlStr + " and i.mwdiv='M'"
'		sqlStr = sqlStr + " group by m.socid"
'		sqlStr = sqlStr + " order by m.socid"

        ''오프샾 매입(10x10 매입) -2011-03-14 수정

        sqlStr = " select m.socid, IsNull(sum(d.sellcash*d.itemno*-1),0) as ttlsellcash,"
        sqlStr = sqlStr + " IsNull(sum(d.suplycash*d.itemno*-1),0) as ttlsuplycash,"
        sqlStr = sqlStr + " IsNull(sum(d.buycash*d.itemno*-1),0) as ttlbuycash"
        sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m"
        sqlStr = sqlStr + "	Join [db_storage].[dbo].tbl_acount_storage_detail d"
        sqlStr = sqlStr + "	on m.code=d.mastercode"
        sqlStr = sqlStr + "	Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "	on d.itemid=i.itemid"
        sqlStr = sqlStr + "	and d.iitemgubun='10'"
        sqlStr = sqlStr + "	and i.mwdiv='M'"
        sqlStr = sqlStr + " 	join  [db_shop].[dbo].tbl_shop_user s"
        sqlStr = sqlStr + "	on m.socid=s.userid"
        sqlStr = sqlStr + " where convert(varchar(7),m.executedt,20)='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + " and m.deldt is NULL"
        sqlStr = sqlStr + " and m.ipchulflag='S'"
        sqlStr = sqlStr + " and d.deldt is NULL"
        sqlStr = sqlStr + " group by m.socid"
        sqlStr = sqlStr + " order by m.socid"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
			set FItemList(i) = new COffShopMonthGainItem
				FItemList(i).FShopid = rsget("socid")
				FItemList(i).FSuplysum = rsget("ttlsuplycash")
				FItemList(i).FTotsum = rsget("ttlbuycash")
			i=i+1
			rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public Sub getOffMonthGainSum_OLD()
		dim sqlStr,i
		dim meaip10x10ttl,meaipupchettl


		''업체 매입 매출데이터
'		sqlStr = " select sum(IsNULL(j.realjungsansum,0)) as realjungsansum, sum(IsNULL(s.realsellsum,0)) as realsellsum"
'		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster j, "
'		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s "
'		sqlStr = sqlStr + " where j.yyyymm='" + FRectYYYYMM + "'"
'		sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
'		sqlStr = sqlStr + " and s.yyyymm='" + FRectYYYYMM + "'"
'		sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "'"
'		sqlStr = sqlStr + " and j.jungsanid=s.makerid"
'		sqlStr = sqlStr + " and j.chargediv='8'"
'
'		rsget.Open sqlStr,dbget,1
'		if Not rsget.Eof then
'			meaipupchettl = rsget("realsellsum")
'		end if
'		rsget.Close


		''정산데이타
		sqlStr = "select IsNULL(j.gubuncd,'') as chargediv, IsNULL(j.comm_name,'') as comm_name, "
		sqlStr = sqlStr + " sum(IsNULL(j.realjungsansum,0)) as realjungsansum, "
		sqlStr = sqlStr + " sum(IsNULL(c.upchebuysum,0)) as upchebuysum, "
		sqlStr = sqlStr + " sum(IsNULL(c.shopsuplysum,0)) as shopsuplysum,"
		sqlStr = sqlStr + " sum(IsNULL(c.chul_upchebuysum,0)) as chul_upchebuysum, "
		sqlStr = sqlStr + " sum(IsNULL(c.chul_shopsuplysum,0)) as chul_shopsuplysum,"
		sqlStr = sqlStr + " sum(IsNULL(c.re_upchebuysum,0)) as re_upchebuysum, "
		sqlStr = sqlStr + " sum(IsNULL(c.re_shopsuplysum,0)) as re_shopsuplysum,"
		sqlStr = sqlStr + " sum(IsNULL(s.realsellsum,0)) as realsellsum"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " on d.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " and u.userid=d.makerid "

        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.gubuncd, c.comm_name, m.makerid,"
        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as realjungsansum"
        sqlStr = sqlStr + "      from "
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d"
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_jungsan_comm_code c on d.gubuncd=c.comm_cd"
        sqlStr = sqlStr + "     where m.idx=d.masteridx"
        sqlStr = sqlStr + "     and m.yyyymm='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + "     and d.shopid='" + FRectShopID + "'"
        sqlStr = sqlStr + "     group by gubuncd, c.comm_name, m.makerid"
        sqlStr = sqlStr + " ) j"
        sqlStr = sqlStr + " on u.userid=j.makerid"

		sqlStr = sqlStr + " left join"
		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c"
		sqlStr = sqlStr + " on  c.yyyymm='" + FRectYYYYMM + "' "
		sqlStr = sqlStr + " and c.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " and u.userid=c.makerid "

		sqlStr = sqlStr + " left join"
		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s"
		sqlStr = sqlStr + " on  s.yyyymm='" + FRectYYYYMM + "' "
		sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " and u.userid=s.makerid "

		sqlStr = sqlStr + " group by IsNULL(j.gubuncd,''), IsNULL(j.comm_name,'')"
		sqlStr = sqlStr + " order by chargediv desc"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof

			set FItemList(i) = new COffShopMonthGainItem
			FItemList(i).FShopID = FRectShopID
			FItemList(i).FChargeDiv = rsget("chargediv")
            FItemList(i).FChargeDivName = rsget("comm_name")
			FItemList(i).Ftotsum = rsget("realsellsum")
			FItemList(i).Frealjungsansum = rsget("realjungsansum")
			FItemList(i).Fupchebuysum	= rsget("upchebuysum")
			FItemList(i).Fshopsuplysum	= rsget("shopsuplysum")
			FItemList(i).Fchul_upchebuysum 	= rsget("chul_upchebuysum")
			FItemList(i).Fchul_shopsuplysum = rsget("chul_shopsuplysum")
			FItemList(i).Fre_upchebuysum 	= rsget("re_upchebuysum")
			FItemList(i).Fre_shopsuplysum 	= rsget("re_shopsuplysum")

            ''업체특정/ 매장매입
			if (FItemList(i).FChargeDiv="B012") or (FItemList(i).FChargeDiv="B022") then
				FItemList(i).Fminuscharge = FItemList(i).Ftotsum - FItemList(i).Frealjungsansum
			''센터에서 오는경우.
			else
				FItemList(i).Fminuscharge = FItemList(i).Ftotsum - FItemList(i).Fshopsuplysum
			end if

			i=i+1
			rsget.moveNext
			loop
		end if
		rsget.Close


        '' Pos 매출/마일리지 사용
		sqlStr = "select sum(IsNULL(realsum,0)) as realsum, sum(IsNULL(spendmile,0)) as spendmile "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and year(m.shopregdate)='" + Left(FRectYYYYMM,4) + "'"
		sqlStr = sqlStr + " and month(m.shopregdate)='" + Right(FRectYYYYMM,2) + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COffShopMonthGainItem
		if  not rsget.EOF  then
			FOneItem.FTotSum = rsget("realsum")
			FOneItem.FTotSpendMile = rsget("spendmile")

			if IsNULL(FOneItem.FTotSum) then FOneItem.FTotSum =0
			if IsNULL(FOneItem.FTotSpendMile) then FOneItem.FTotSpendMile =0
		end if
		rsget.Close
	end sub

    public Sub getOffMonthGainSum()
		dim sqlStr,i
		dim meaip10x10ttl,meaipupchettl

        sqlStr = " select T1.*, j.* from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " select IsNULL(d.comm_cd,'') as chargediv, IsNULL(cd.comm_name,'') as comm_name, "
        sqlStr = sqlStr + " sum(IsNULL(c.upchebuysum,0)) as upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.shopsuplysum,0)) as shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(c.chul_upchebuysum,0)) as chul_upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.chul_shopsuplysum,0)) as chul_shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(c.re_upchebuysum,0)) as re_upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.re_shopsuplysum,0)) as re_shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(s.realsellsum,0)) as realsellsum"
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u"

        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
        sqlStr = sqlStr + " on d.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=d.makerid "

        sqlStr = sqlStr + " left join"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s"
        sqlStr = sqlStr + " on  s.yyyymm='" + FRectYYYYMM + "' "
        sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=s.makerid "


        sqlStr = sqlStr + " left join"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c"
        sqlStr = sqlStr + " on  c.yyyymm='" + FRectYYYYMM + "' "
        sqlStr = sqlStr + " and c.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=c.makerid "

        sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_jungsan_comm_code cd on IsNULL(d.comm_cd,'')=cd.comm_cd"

        sqlStr = sqlStr + " group by IsNULL(d.comm_cd,''), IsNULL(cd.comm_name,'')"
        sqlStr = sqlStr + " ) T1"


        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.gubuncd, "
        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as realjungsansum"
        sqlStr = sqlStr + "      from "
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d"
        sqlStr = sqlStr + "     where m.idx=d.masteridx"
        sqlStr = sqlStr + "     and m.yyyymm='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + "     and d.shopid='" + FRectShopID + "'"
        sqlStr = sqlStr + "     group by gubuncd"
        sqlStr = sqlStr + " ) j"
        sqlStr = sqlStr + " on T1.chargediv=j.gubuncd"
        sqlStr = sqlStr + " order by T1.chargediv"
'response.write sqlStr

'		sqlStr = "select IsNULL(j.gubuncd,d.comm_cd) as chargediv, IsNULL(cd.comm_name,'') as comm_name, "
'		sqlStr = sqlStr + " sum(IsNULL(j.realjungsansum,0)) as realjungsansum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.upchebuysum,0)) as upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.shopsuplysum,0)) as shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(c.chul_upchebuysum,0)) as chul_upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.chul_shopsuplysum,0)) as chul_shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(c.re_upchebuysum,0)) as re_upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.re_shopsuplysum,0)) as re_shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(s.realsellsum,0)) as realsellsum"
'		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u"
'
'        sqlStr = sqlStr + " left join"
'		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s"
'		sqlStr = sqlStr + " on  s.yyyymm='" + FRectYYYYMM + "' "
'		sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=s.makerid "
'
'		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
'		sqlStr = sqlStr + " on d.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=d.makerid "
'
'        sqlStr = sqlStr + " left join ("
'        sqlStr = sqlStr + "     select d.gubuncd, m.makerid,"
'        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as realjungsansum"
'        sqlStr = sqlStr + "      from "
'        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m,"
'        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d"
'        sqlStr = sqlStr + "     where m.idx=d.masteridx"
'        sqlStr = sqlStr + "     and m.yyyymm='" + FRectYYYYMM + "'"
'        sqlStr = sqlStr + "     and d.shopid='" + FRectShopID + "'"
'        sqlStr = sqlStr + "     group by gubuncd, m.makerid"
'        sqlStr = sqlStr + " ) j"
'        sqlStr = sqlStr + " on u.userid=j.makerid"
'
'        sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_jungsan_comm_code cd on IsNULL(j.gubuncd,d.comm_cd)=cd.comm_cd"
'
'		sqlStr = sqlStr + " left join"
'		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c"
'		sqlStr = sqlStr + " on  c.yyyymm='" + FRectYYYYMM + "' "
'		sqlStr = sqlStr + " and c.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=c.makerid "
'
'		sqlStr = sqlStr + " group by IsNULL(j.gubuncd,d.comm_cd), IsNULL(cd.comm_name,'')"
'		sqlStr = sqlStr + " order by chargediv desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof

			set FItemList(i) = new COffShopMonthGainItem
			FItemList(i).FShopID = FRectShopID
			FItemList(i).FChargeDiv         = Trim(rsget("chargediv"))
            FItemList(i).FChargeDivName     = rsget("comm_name")
			FItemList(i).Ftotsum            = rsget("realsellsum")
			FItemList(i).Frealjungsansum    = rsget("realjungsansum")
			if IsNULL(FItemList(i).Frealjungsansum) then FItemList(i).Frealjungsansum = 0
			FItemList(i).Fupchebuysum	    = rsget("upchebuysum")
			FItemList(i).Fshopsuplysum	    = rsget("shopsuplysum")
			FItemList(i).Fchul_upchebuysum 	= rsget("chul_upchebuysum")
			FItemList(i).Fchul_shopsuplysum = rsget("chul_shopsuplysum")
			FItemList(i).Fre_upchebuysum 	= rsget("re_upchebuysum")
			FItemList(i).Fre_shopsuplysum 	= rsget("re_shopsuplysum")

            ''업체특정/ 매장매입
			if (FItemList(i).FChargeDiv="B012") or (FItemList(i).FChargeDiv="B022") then
				FItemList(i).Fminuscharge   = FItemList(i).Ftotsum - FItemList(i).Frealjungsansum
			''센터에서 오는경우
			else
				FItemList(i).Fminuscharge   = FItemList(i).Ftotsum - FItemList(i).Fshopsuplysum
			end if

			i=i+1
			rsget.moveNext
			loop
		end if
		rsget.Close


        '' Pos 매출/마일리지 사용
		sqlStr = "select sum(IsNULL(realsum,0)) as realsum, sum(IsNULL(spendmile,0)) as spendmile "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and year(m.shopregdate)='" + Left(FRectYYYYMM,4) + "'"
		sqlStr = sqlStr + " and month(m.shopregdate)='" + Right(FRectYYYYMM,2) + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COffShopMonthGainItem
		if  not rsget.EOF  then
			FOneItem.FTotSum = rsget("realsum")
			FOneItem.FTotSpendMile = rsget("spendmile")

			if IsNULL(FOneItem.FTotSum) then FOneItem.FTotSum =0
			if IsNULL(FOneItem.FTotSpendMile) then FOneItem.FTotSpendMile =0
		end if
		rsget.Close
	end sub


	public Sub getFrnMonthGainSum()
		dim sqlStr,i
		dim upchewitakmeachul, meaipchulgomeachul


		''가맹점 법인 매출 데이터
		sqlStr = " select m.divcode, sum(s.totalsuplycash) as meachul"
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_master m,"
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_submaster s"
        sqlStr = sqlStr + " where m.idx=s.masteridx"
        sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
        sqlStr = sqlStr + " and convert(varchar(7),s.execdate,21)='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + " and m.statecd>0"
        sqlStr = sqlStr + " group by m.divcode"
        sqlStr = sqlStr + " order by m.divcode"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
		    do until rsget.eof
		        if rsget("divcode")="WS" then
			        upchewitakmeachul = rsget("meachul")
			    elseif rsget("divcode")="MC" then
			        meaipchulgomeachul = rsget("meachul")
			    end if
			    i=i+1
			    rsget.moveNext

			loop
		end if
		rsget.Close

        sqlStr = " select T1.*, j.* from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " select IsNULL(d.comm_cd,'') as chargediv, IsNULL(cd.comm_name,'') as comm_name, "
        sqlStr = sqlStr + " sum(IsNULL(c.upchebuysum,0)) as upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.shopsuplysum,0)) as shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(c.chul_upchebuysum,0)) as chul_upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.chul_shopsuplysum,0)) as chul_shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(c.re_upchebuysum,0)) as re_upchebuysum, "
        sqlStr = sqlStr + " sum(IsNULL(c.re_shopsuplysum,0)) as re_shopsuplysum,"
        sqlStr = sqlStr + " sum(IsNULL(s.realsellsum,0)) as realsellsum"
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u"

        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
        sqlStr = sqlStr + " on d.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=d.makerid "

        sqlStr = sqlStr + " left join"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s"
        sqlStr = sqlStr + " on  s.yyyymm='" + FRectYYYYMM + "' "
        sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=s.makerid "


        sqlStr = sqlStr + " left join"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c"
        sqlStr = sqlStr + " on  c.yyyymm='" + FRectYYYYMM + "' "
        sqlStr = sqlStr + " and c.shopid='" + FRectShopID + "' "
        sqlStr = sqlStr + " and u.userid=c.makerid "

        sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_jungsan_comm_code cd on IsNULL(d.comm_cd,'')=cd.comm_cd"

        sqlStr = sqlStr + " group by IsNULL(d.comm_cd,''), IsNULL(cd.comm_name,'')"
        sqlStr = sqlStr + " ) T1"


        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.gubuncd, "
        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as realjungsansum"
        sqlStr = sqlStr + "      from "
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d"
        sqlStr = sqlStr + "     where m.idx=d.masteridx"
        sqlStr = sqlStr + "     and m.yyyymm='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + "     and d.shopid='" + FRectShopID + "'"
        sqlStr = sqlStr + "     group by gubuncd"
        sqlStr = sqlStr + " ) j"
        sqlStr = sqlStr + " on T1.chargediv=j.gubuncd"
        sqlStr = sqlStr + " order by T1.chargediv"

		''정산데이타
'		sqlStr = "select IsNULL(j.gubuncd,d.comm_cd) as chargediv, IsNULL(cd.comm_name,'') as comm_name, "
'		sqlStr = sqlStr + " sum(IsNULL(j.realjungsansum,0)) as realjungsansum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.upchebuysum,0)) as upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.shopsuplysum,0)) as shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(c.chul_upchebuysum,0)) as chul_upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.chul_shopsuplysum,0)) as chul_shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(c.re_upchebuysum,0)) as re_upchebuysum, "
'		sqlStr = sqlStr + " sum(IsNULL(c.re_shopsuplysum,0)) as re_shopsuplysum,"
'		sqlStr = sqlStr + " sum(IsNULL(s.realsellsum,0)) as realsellsum"
'		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u"
'
'		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
'		sqlStr = sqlStr + " on d.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=d.makerid "
'
'		sqlStr = sqlStr + " left join ("
'        sqlStr = sqlStr + "     select d.gubuncd, m.makerid,"
'        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as realjungsansum"
'        sqlStr = sqlStr + "      from "
'        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m,"
'        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d"
'        sqlStr = sqlStr + "     where m.idx=d.masteridx"
'        sqlStr = sqlStr + "     and m.yyyymm='" + FRectYYYYMM + "'"
'        sqlStr = sqlStr + "     and d.shopid='" + FRectShopID + "'"
'        sqlStr = sqlStr + "     group by gubuncd, m.makerid"
'        sqlStr = sqlStr + " ) j"
'        sqlStr = sqlStr + " on u.userid=j.makerid"
'
'        sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_jungsan_comm_code cd on IsNULL(j.gubuncd,d.comm_cd)=cd.comm_cd"
'
'		sqlStr = sqlStr + " left join"
'		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c"
'		sqlStr = sqlStr + " on  c.yyyymm='" + FRectYYYYMM + "' "
'		sqlStr = sqlStr + " and c.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=c.makerid "
'
'		sqlStr = sqlStr + " left join"
'		sqlStr = sqlStr + " [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s"
'		sqlStr = sqlStr + " on  s.yyyymm='" + FRectYYYYMM + "' "
'		sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "' "
'		sqlStr = sqlStr + " and u.userid=s.makerid "
'
'		sqlStr = sqlStr + " group by IsNULL(j.gubuncd,d.comm_cd), IsNULL(cd.comm_name,'')"
'		sqlStr = sqlStr + " order by chargediv desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof

			set FItemList(i) = new COffShopMonthGainItem
			FItemList(i).FShopID = FRectShopID
			FItemList(i).FChargeDiv         = trim(rsget("chargediv"))
            FItemList(i).FChargeDivName     = rsget("comm_name")

			FItemList(i).Ftotsum            = rsget("realsellsum")
			FItemList(i).Frealjungsansum    = rsget("realjungsansum")
			if IsNULL(FItemList(i).Frealjungsansum) then FItemList(i).Frealjungsansum=0

			FItemList(i).Fupchebuysum	    = rsget("upchebuysum")
			FItemList(i).Fshopsuplysum	    = rsget("shopsuplysum")
			FItemList(i).Fchul_upchebuysum 	= rsget("chul_upchebuysum")
			FItemList(i).Fchul_shopsuplysum = rsget("chul_shopsuplysum")
			FItemList(i).Fre_upchebuysum 	= rsget("re_upchebuysum")
			FItemList(i).Fre_shopsuplysum 	= rsget("re_shopsuplysum")

            ''//매장매입은 없음.
			if (FItemList(i).FChargeDiv="B012") then
			    FItemList(i).Fshopsuplysum = upchewitakmeachul
			end if

			FItemList(i).Fminuscharge = FItemList(i).Ftotsum - FItemList(i).Fshopsuplysum
			i=i+1
			rsget.moveNext
			loop
		end if
		rsget.Close



		sqlStr = "select sum(IsNULL(realsum,0)) as realsum, sum(IsNULL(spendmile,0)) as spendmile "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and year(m.shopregdate)='" + Left(FRectYYYYMM,4) + "'"
		sqlStr = sqlStr + " and month(m.shopregdate)='" + Right(FRectYYYYMM,2) + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COffShopMonthGainItem
		if  not rsget.EOF  then
			FOneItem.FTotSum = rsget("realsum")
			FOneItem.FTotSpendMile = rsget("spendmile")

			if IsNULL(FOneItem.FTotSum) then FOneItem.FTotSum =0
			if IsNULL(FOneItem.FTotSpendMile) then FOneItem.FTotSpendMile =0
		end if
		rsget.Close
	end sub




	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>