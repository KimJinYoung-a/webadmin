<%
class COffshopJungsanItem
	public Fidx
	public Fyyyymm
	public Fshopid
	public Fjungsanid
	public Ftotitemcnt
	public Ftotsum
	public Fminuscharge
	public Fchargepercent
	public Frealjungsansum
	public Fbigo
	public Fcurrstate
	public Fsegumil
	public Fipkumil
	public Fregdate
	public Fchargediv
	public Ffranchargediv
	public Fgroupidx
	public Ftaxregdate
	public Fdifferencekey
	public Ftaxtype
	public Ftaxlinkidx
	public Fneotaxno
	public Foffgubun

	public FSelltotitemcnt
	public FSelltotsum

	public FOrgChargeDiv
	public Fjungsan_acctname
	public Fjungsan_bank
	public Fjungsan_acctno
	public Fcompany_name
	public Fjungsan_gubun
	public Fcompany_no
	public FJungsanChargediv

	public FNoFixSum
	public FFixSum
	public FIpkumSum
	public FAutoJungsan

	public FShopName
	public FShopDiv
	public Fdefaultmargin
	public Fdefaultsuplymargin

	public FFixsegumil

	public FOffMeaip
	public FOnMeaip
	public FonWitak
	public FOffChulgo
	public FOffSell

	public Fonlinedefaultmargine
	public Fonlinemaeipdiv

	public Fjungsan_date_off


	public function IsMaeipJungsan()
		IsMaeipJungsan = (Fonlinemaeipdiv="M") and (FOrgChargeDiv="4" or FOrgChargeDiv="5")
	end function

	public function GetOnlineMaeipDivName
		if Fonlinemaeipdiv="M" then
			GetOnlineMaeipDivName = "매입"
		elseif Fonlinemaeipdiv="W" then
			GetOnlineMaeipDivName = "위탁"
		elseif Fonlinemaeipdiv="U" then
			GetOnlineMaeipDivName = "업체"
		else
			GetOnlineMaeipDivName = Fonlinemaeipdiv
		end if

	end function


	public function GetTotalTaxSuply()
		if Ftaxtype="01" then
			GetTotalTaxSuply = CLng(Frealjungsansum / 1.1)
		else
			GetTotalTaxSuply = Frealjungsansum
		end if
	end function

	public function GetTotalTaxVat()
		GetTotalTaxVat = Frealjungsansum - GetTotalTaxSuply
	end function

	public function getDbDate()
		dim sqlstr
		sqlstr = " select convert(varchar(10),getdate(),21) as nowdate "
		rsget.Open sqlStr,dbget,1
		getDbDate = CDate(rsget("nowdate"))
		rsget.Close
	end function

	public function GetNormalTaxDate()
		if Not(IsNULL(FFixsegumil)) and (FFixsegumil<>"") then
			GetNormalTaxDate = FFixsegumil
		else
			GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1)
		end if
	end function

	public function GetPreFixSegumil()
		dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()
        
        '' 12일까지 마감할 경우 13으로 세팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"13")
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"13")

		if (thisdate>=ithis21day) then
			GetPreFixSegumil = ithis1day
		elseif (maytaxdate<premonth21day) then
			GetPreFixSegumil = premonth1day
		else
			GetPreFixSegumil = maytaxdate
		end if
		
		
		
		
	end function

	'public function GetPreFixSegumil()
	'	dim thisdate, maytaxdate
	'	dim i0116, i0416, i0716, i1016
	'	dim i0101, i0401, i0701, i1001

	'	thisdate = getDbDate()
	'	maytaxdate = GetNormalTaxDate()

	'	i0116 = dateserial(Left(thisdate,4),"01","16")
	'	i0416 = dateserial(Left(thisdate,4),"04","16")
	'	i0716 = dateserial(Left(thisdate,4),"07","16")
	'	i1016 = dateserial(Left(thisdate,4),"10","16")

	'	i0101 = dateserial(Left(thisdate,4),"01","01")
	'	i0401 = dateserial(Left(thisdate,4),"04","01")
	'	i0701 = dateserial(Left(thisdate,4),"07","01")
	'	i1001 = dateserial(Left(thisdate,4),"10","01")

	'	if ((thisdate>=i1016) and (maytaxdate<i1001)) then
	'		GetPreFixSegumil = i1001
	'	elseif ((thisdate>=i0716) and (maytaxdate<i0701)) then
	'		GetPreFixSegumil = i0701
	'	elseif ((thisdate>=i0416) and (maytaxdate<i0401)) then
	'		GetPreFixSegumil = i0401
	'	elseif ((thisdate>=i0116) and (maytaxdate<i0101)) then
	'		GetPreFixSegumil = i0101
	'	else
	'		GetPreFixSegumil = maytaxdate
	'	end if
	'end function


	public function IsElecTaxExists()
		IsElecTaxExists = Not(IsNULL(FTaxLinkidx) or (FTaxLinkidx="")) and (Fcurrstate>=3)
	end function


	''//세금계산서
	public function IsElecTaxCase()
		IsElecTaxCase = (Ftaxtype="01") and (Fjungsan_gubun="일반과세") and (Fcurrstate<3)
	end function


	''//계산서
	public function IsElecFreeTaxCase()
		IsElecFreeTaxCase = (Ftaxtype="02") 'and (Fjungsan_gubun="면세")
	end function


	''//간이, 원천, 기타
	public function IsElecSimpleBillCase()
		IsElecSimpleBillCase = (Ftaxtype="03") and (Fcurrstate<3)
	end function

	public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "간이"
		end if
	end function

	public function GetTaxtypeNameColor()
		if Ftaxtype="01" then
			GetTaxtypeNameColor = "#000000"
		elseif Ftaxtype="02" then
			GetTaxtypeNameColor = "#FF3333"
		elseif Ftaxtype="03" then
			GetTaxtypeNameColor = "#3333FF"
		end if
	end function

	public function GetShopDivName()
		if IsNull(FShopDiv) then

		elseif FShopDiv="1" then
			GetShopDivName = "직영"
		elseif FShopDiv="2" then
			GetShopDivName = "수수료매장"
		elseif FShopDiv="3" then
			GetShopDivName = "가맹점"
		end if
	end function


	public function GetAutoJungsanName()
		if IsNull(FAutoJungsan) then

		elseif FAutoJungsan="Y" then
			GetAutoJungsanName = "자동"
		elseif FAutoJungsan="N" then
			GetAutoJungsanName = "수기"
		end if
	end function

	public function GetAutoJungsanColor()
		if IsNull(FAutoJungsan) then
			GetAutoJungsanColor = "#000000"
		elseif FAutoJungsan="Y" then
			GetAutoJungsanColor = "#000000"
		elseif FAutoJungsan="N" then
			GetAutoJungsanColor = "#4444AA"
		end if
	end function

	public function GetCurrStateName()
		if IsNull(Fcurrstate) or (Fcurrstate="") then
			GetCurrStateName = "미정산"
		elseif Fcurrstate="0" then
			GetCurrStateName = "수정중"
		elseif Fcurrstate="1" then
			GetCurrStateName = "업체확인중"
		elseif Fcurrstate="2" then
			GetCurrStateName = "업체확인완료"
		elseif Fcurrstate="3" then
			GetCurrStateName = "정산확정"
		elseif Fcurrstate="7" then
			GetCurrStateName = "입금완료"
		elseif Fcurrstate="8" then
			GetCurrStateName = "정산안함"
		elseif Fcurrstate="9" then
			GetCurrStateName = "통합정산"
		end if
	end function

	public function GetStateColor()
		if Fcurrstate="0" then
			GetStateColor = "#000000"
		elseif Fcurrstate="1" then
			GetStateColor = "#448888"
		elseif Fcurrstate="2" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="3" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="7" then
			GetStateColor = "#FF0000"
		elseif Fcurrstate="8" then
			GetStateColor = "#AAAAAA"
		elseif Fcurrstate="9" then
			GetStateColor = "#888844"
		else

		end if
	end function

	public function getOrgChargeDivName()
		if FOrgChargeDiv="2" then
			getOrgChargeDivName = "10x10 위탁"
		elseif FOrgChargeDiv="4" then
			getOrgChargeDivName = "10x10 매입"
		elseif FOrgChargeDiv="5" then
			getOrgChargeDivName = "출고분정산"
		elseif FOrgChargeDiv="6" then
			getOrgChargeDivName = "업체 위탁"
		elseif FOrgChargeDiv="8" then
			getOrgChargeDivName = "업체 매입"
		elseif FOrgChargeDiv="9" then
			getOrgChargeDivName = "가맹점"
		elseif FOrgChargeDiv="0" then
			getOrgChargeDivName = "통합"
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 위탁"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 위탁"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		end if
	end function

	public function getJungSanChargeDivNameUpcheView()
		if FJungsanChargediv="2" then
			getJungSanChargeDivNameUpcheView = "위탁"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivNameUpcheView = "위탁"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivNameUpcheView = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivNameUpcheView = "통합"
		end if
	end function

	public function GetFranChargeDivName()
		if FFranChargeDiv="2" then
			GetFranChargeDivName = "10x10위탁"
		elseif FFranChargeDiv="4" then
			GetFranChargeDivName = "10x10매입"
		elseif FFranChargeDiv="5" then
			GetFranChargeDivName = "출고분정산"
		elseif FFranChargeDiv="6" then
			GetFranChargeDivName = "업체위탁"
		elseif FFranChargeDiv="8" then
			GetFranChargeDivName = "업체매입"
		end if
	end function

	public function GetOffgubunName()
		if Foffgubun="OFF" then
			GetOffgubunName = "오프라인"
		elseif Foffgubun="FRN" then
			GetOffgubunName = "가맹점"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end class


Class COffShopJungSanDetailItem
	public Fidx
	public Fmasteridx
	public Forderno
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Frealsellprice
	public Fsuplyprice
	public Fitemno
	public Fmakerid
	public Flinkidx
	public Fjungsangubun

	public Fonlinemwdiv
	public Fshopid

	public function getDetailGubunName()
		if Fjungsangubun = "101" then
			getDetailGubunName = "매입"
		elseif Fjungsangubun = "131" then
			getDetailGubunName = "위탁재고->매입"
		elseif Fjungsangubun = "111" then
			getDetailGubunName = "위탁"
		elseif Fjungsangubun = "121" then
			getDetailGubunName = "위탁"
		elseif Fjungsangubun = "801" then
			getDetailGubunName = "off매입"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


class CDetailGroupSubItem
	public Fidx
	public Fgroupidx
	public Fyyyymm
	public Fshopid
	public Fshopname

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class


class COffshopJungsan
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectJungsanId
	public FRectShopID
	public FRectOnlyUpcheJungSan
	public FRectOnlyFranUpcheJungSan

	public FRectOnlyOwnOffJungsan
	public FRectOnlyFranOffJungsan

	public FRectOnlyUpcheview

	public FRectYYYYMM
	public FRectIdx
	public FRectGroupidx

	public FRectchargediv
	public FRectCurrState

	public FRectStartDate
	public FRectEndDate

	public FRectMwgubun
	public FRectonoffgubun

	public Sub GetOffJungSanBatchList()
		dim i,sqlStr
		dim nextYYYYMMDD

		nextYYYYMMDD = CStr(dateserial(Left(FRectYYYYMM,4),Right(FRectYYYYMM,2)+1,1))

		if (FRectChargediv="2") or (FRectChargediv="6") then
			''위탁, 업체위탁 : 판매분정산
			sqlStr = " select T.*, u.chargediv,u.autojungsan, u.defaultmargin, u.defaultsuplymargin,"
			sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
			sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
			sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
			sqlStr = sqlStr + " IsNULL(j.minuscharge,0) as minuscharge,"
			sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum,"
			sqlStr = sqlStr + " j.chargediv as jchargediv, "
			sqlStr = sqlStr + " c.maeipdiv as onlinemaeipdiv, c.defaultmargine as onlinedefaultmargine"

			sqlStr = sqlStr + " from ("
			''판매내역.
			sqlStr = sqlStr + " 	select m.shopid, d.makerid, "
			sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

			sqlStr = sqlStr + " 	where m.idx=d.masteridx"
			sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectYYYYMM + "-" + "01'"
			sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " 	and m.cancelyn='N'"
			sqlStr = sqlStr + " 	and d.cancelyn='N'"
			sqlStr = sqlStr + " 	group by m.shopid,d.makerid"
			sqlStr = sqlStr + " ) as T"

			sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
			sqlStr = sqlStr + " on j.makerid=T.makerid"
			sqlStr = sqlStr + " and j.shopid=T.shopid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"

			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
			sqlStr = sqlStr + " on u.makerid=T.makerid and u.shopid=T.shopid "

			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
			sqlStr = sqlStr + " on u.makerid=c.userid "


			sqlStr = sqlStr + " where u.chargediv='" + FRectChargeDiv + "'"
			sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"


		elseif (FRectChargediv="5") or (FRectChargediv="4") then
			''매입출고, 텐바이텐매입
			sqlStr = " select T.*, u.chargediv, u.autojungsan, u.defaultmargin, u.defaultsuplymargin,"
			sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
			sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
			sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
			sqlStr = sqlStr + " IsNULL(j.minuscharge,0) as minuscharge,"
			sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum,"
			sqlStr = sqlStr + " j.chargediv as jchargediv,"
			sqlStr = sqlStr + " c.maeipdiv as onlinemaeipdiv, c.defaultmargine as onlinedefaultmargine"

			sqlStr = sqlStr + " from ("
			''출고내역.
			sqlStr = sqlStr + " 	select m.socid as shopid, d.imakerid as makerid , "
			sqlStr = sqlStr + " 	count(d.itemno) as totno, sum(d.sellcash*d.itemno*-1) as totsum"
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
			sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
			sqlStr = sqlStr + " 	where m.code=d.mastercode"
			sqlStr = sqlStr + " 	and m.executedt>='" + FRectYYYYMM + "-" + "01'"
			sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " 	and m.socid='" + FRectShopID + "'"
			sqlStr = sqlStr + " 	and m.deldt is null"
			sqlStr = sqlStr + " 	and d.deldt  is null"
			sqlStr = sqlStr + " 	group by m.socid, d.imakerid"

			sqlStr = sqlStr + " ) as T"

			sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
			sqlStr = sqlStr + " on j.makerid=T.makerid"
			sqlStr = sqlStr + " and j.shopid=T.shopid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"

			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
			sqlStr = sqlStr + " on u.makerid=T.makerid and u.shopid=T.shopid "

			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
			sqlStr = sqlStr + " on u.makerid=c.userid "

			sqlStr = sqlStr + " where u.chargediv='" + FRectChargeDiv + "'"
			sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"
		elseif (FRectChargediv="8")  then
			''업체매입
			sqlStr = " select T.*, u.chargediv, u.autojungsan, u.defaultmargin, u.defaultsuplymargin,"
			sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
			sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
			sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
			sqlStr = sqlStr + " IsNULL(j.minuscharge,0) as minuscharge,"
			sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum,"
			sqlStr = sqlStr + " j.chargediv as jchargediv,"
			sqlStr = sqlStr + " c.maeipdiv as onlinemaeipdiv, c.defaultmargine as onlinedefaultmargine"

			sqlStr = sqlStr + " from ("
			''업체직접입고내역
			sqlStr = sqlStr + " 	select m.shopid as shopid, m.chargeid as makerid , "
			sqlStr = sqlStr + " 	count(d.itemno) as totno, sum(d.sellcash*d.itemno) as totsum"
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_master m,"
			sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " 	where m.idx=d.masteridx"
			sqlStr = sqlStr + " 	and m.execdt>='" + FRectYYYYMM + "-" + "01'"
			sqlStr = sqlStr + " 	and m.execdt<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " 	and m.statecd>=7"
			sqlStr = sqlStr + " 	and m.deleteyn='N'"
			sqlStr = sqlStr + " 	and d.deleteyn='N'"
			sqlStr = sqlStr + " 	group by m.shopid, m.chargeid"

			sqlStr = sqlStr + " ) as T"

			sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
			sqlStr = sqlStr + " on j.makerid=T.makerid"
			sqlStr = sqlStr + " and j.shopid=T.shopid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"

			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
			sqlStr = sqlStr + " on u.makerid=T.makerid and u.shopid=T.shopid "

			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
			sqlStr = sqlStr + " on u.makerid=c.userid "

			sqlStr = sqlStr + " where u.chargediv='" + FRectChargeDiv + "'"
			sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"
		end if


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fjungsanid = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectYYYYMM

					''판매분,입출고분
					FItemList(i).FSelltotitemcnt = rsget("totno")
					FItemList(i).FSelltotsum     = rsget("totsum")

					''정산내역분
					FItemList(i).Ftotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Ftotsum     = rsget("jungsantotsum")
					FItemList(i).FIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fdefaultmargin		= rsget("defaultmargin")
					FItemList(i).Fdefaultsuplymargin= rsget("defaultsuplymargin")

					''정산시정산구분
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					''현재정산구분
					FItemList(i).FOrgChargeDiv = rsget("chargediv")

					FItemList(i).FAutojungsan = rsget("autojungsan")

					FItemList(i).Fonlinedefaultmargine 	= rsget("onlinedefaultmargine")
					FItemList(i).Fonlinemaeipdiv		= rsget("onlinemaeipdiv")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetOffJungSanDetailByChulgoShopSum()
		dim i,sqlStr
		dim isEof

		sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		if FRectJungsanId<>"" then
				sqlStr = sqlStr + " and jungsanid='" + CStr(FRectJungsanId) + "'"
		end if

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			set FOneItem = new COffShopJungsanItem

			FOneItem.FShopid	 = rsget("shopid")
			FOneItem.Fjungsanid = rsget("jungsanid")
			FOneItem.FYYYYMM     = rsget("yyyymm")
			FOneItem.Ftotitemcnt      = rsget("totitemcnt")
			FOneItem.Ftotsum        = rsget("totsum")
			FOneItem.Fcurrstate = rsget("currstate")
			FOneItem.FminusCharge    = rsget("minuscharge")
			FOneItem.FRealjungsansum = rsget("realjungsansum")

			FOneItem.Fjungsanchargediv = rsget("chargediv")

			FOneItem.FSegumil = rsget("segumil")
			FOneItem.Fipkumil = rsget("ipkumil")
			'FOneItem.Fregdate = rsget("regdate")
		else
			isEof = true
		end if

		rsget.Close

		if isEof then dbget.close()	:	response.End

		sqlStr = "select j.makerid, j.jungsangubun, s.socid,"
		sqlStr = sqlStr + " sum(j.sellprice*j.itemno) as sellpricesum, "
		sqlStr = sqlStr + " sum(j.realsellprice*j.itemno) as realsellpricesum,"
		sqlStr = sqlStr + " sum(j.suplyprice*j.itemno) as suplypricesum, "
		sqlStr = sqlStr + " sum(j.itemno) as itemnosum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsandetail j"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_acount_storage_master s"
		sqlStr = sqlStr + " on j.orderno=s.code"
		sqlStr = sqlStr + " and s.deldt is null"
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " group by j.makerid, j.jungsangubun,  s.socid"
		sqlStr = sqlStr + " order by s.socid "

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					'FItemList(i).Fmasteridx      = rsget("masteridx")
					'FItemList(i).Forderno        = rsget("orderno")
					FItemList(i).Fsellprice      = rsget("sellpricesum")
					FItemList(i).Frealsellprice  = rsget("realsellpricesum")
					FItemList(i).Fsuplyprice     = rsget("suplypricesum")
					FItemList(i).Fitemno        = rsget("itemnosum")
					FItemList(i).Fmakerid        = rsget("makerid")
					'FItemList(i).Flinkidx        = rsget("linkidx")
					FItemList(i).Fjungsangubun   = rsget("jungsangubun")
					FItemList(i).Fshopid   = rsget("socid")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end Sub

	public Sub GetOffJungSanDetailByChulgoSum()
		dim i,sqlStr
		dim isEof

		sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		if FRectJungsanId<>"" then
				sqlStr = sqlStr + " and jungsanid='" + CStr(FRectJungsanId) + "'"
		end if

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			set FOneItem = new COffShopJungsanItem

			FOneItem.FShopid	 = rsget("shopid")
			FOneItem.Fjungsanid = rsget("jungsanid")
			FOneItem.FYYYYMM     = rsget("yyyymm")
			FOneItem.Ftotitemcnt      = rsget("totitemcnt")
			FOneItem.Ftotsum        = rsget("totsum")
			FOneItem.Fcurrstate = rsget("currstate")
			FOneItem.FminusCharge    = rsget("minuscharge")
			FOneItem.FRealjungsansum = rsget("realjungsansum")

			FOneItem.Fjungsanchargediv = rsget("chargediv")

			FOneItem.FSegumil = rsget("segumil")
			FOneItem.Fipkumil = rsget("ipkumil")
			'FOneItem.Fregdate = rsget("regdate")
		else
			isEof = true
		end if

		rsget.Close

		if isEof then dbget.close()	:	response.End

		sqlStr = "select j.masteridx, j.orderno, j.makerid, j.jungsangubun, s.socid,"
		sqlStr = sqlStr + " sum(j.sellprice*j.itemno) as sellpricesum, "
		sqlStr = sqlStr + " sum(j.realsellprice*j.itemno) as realsellpricesum,"
		sqlStr = sqlStr + " sum(j.suplyprice*j.itemno) as suplypricesum, "
		sqlStr = sqlStr + " sum(j.itemno) as itemnosum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsandetail j"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_acount_storage_master s"
		sqlStr = sqlStr + " on j.orderno=s.code"
		sqlStr = sqlStr + " and s.deldt is null"
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " group by j.masteridx,j.orderno,j.makerid, j.jungsangubun,  s.socid"
		sqlStr = sqlStr + " order by j.orderno "

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					FItemList(i).Fmasteridx      = rsget("masteridx")
					FItemList(i).Forderno        = rsget("orderno")
					FItemList(i).Fsellprice      = rsget("sellpricesum")
					FItemList(i).Frealsellprice  = rsget("realsellpricesum")
					FItemList(i).Fsuplyprice     = rsget("suplypricesum")
					FItemList(i).Fitemno        = rsget("itemnosum")
					FItemList(i).Fmakerid        = rsget("makerid")
					'FItemList(i).Flinkidx        = rsget("linkidx")
					FItemList(i).Fjungsangubun   = rsget("jungsangubun")
					FItemList(i).Fshopid   = rsget("socid")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end Sub

	public sub GetjungsanWithChulgo
		dim i,sqlStr

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " d.shopid, d.makerid,"
		sqlStr = sqlStr + " d.defaultmargin, d.defaultsuplymargin, d.chargediv as orgchargediv,"
		sqlStr = sqlStr + " m.idx, m.yyyymm, m.jungsanid, m.realjungsansum, m.chargediv, m.currstate, "
		sqlStr = sqlStr + " m.franchargediv, m.offgubun, "
		sqlStr = sqlStr + " s.shopname, IsNULL(MI.offmeaip,0) as offmeaip ,IsNULL(MC.offchulgosum,0) as offchulgo "
		sqlStr = sqlStr + " ,IsNULL(MS.realsellsum,0) as offsell "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster m on d.shopid=m.shopid and d.makerid=m.jungsanid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on d.shopid=s.userid"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select socid,divcode,"
		sqlStr = sqlStr + " 	sum(totalsuplycash) as offmeaip "
		sqlStr = sqlStr + " 	from [db_storage].[dbo].tbl_acount_storage_master "
		sqlStr = sqlStr + " 	where executedt>='" + FRectStartDate + "' "
		sqlStr = sqlStr + " 	and executedt<'" + FRectEndDate + "' "
		sqlStr = sqlStr + " 	and deldt is null "
		sqlStr = sqlStr + " 	and divcode='801' "
		sqlStr = sqlStr + " 	group by socid,divcode "
		sqlStr = sqlStr + " ) MI on d.shopid='streetshop800' and d.makerid=MI.socid "

		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select (Left(m.socid,11) + '00') as shopid,d.imakerid,"
		sqlStr = sqlStr + " 	sum(d.buycash*d.itemno*-1) as offchulgosum"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectStartDate + "' "
		sqlStr = sqlStr + " 	and m.executedt<'" + FRectEndDate + "' "
		sqlStr = sqlStr + " 	and m.deldt is null "
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.divcode='006' "
		sqlStr = sqlStr + " 	group by (Left(m.socid,11) + '00') ,d.imakerid"
		sqlStr = sqlStr + " ) MC on d.shopid=MC.shopid and d.makerid=MC.imakerid "

		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select (Left(shopid,11) + '00') as shopid, makerid,"
		sqlStr = sqlStr + " 	sum(offchulgosum) as offchulgosum, sum(realsellsum) as realsellsum "
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shop_meachul_summary "
		sqlStr = sqlStr + " 	where yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " 	group by (Left(shopid,11) + '00'), makerid"
		sqlStr = sqlStr + " ) MS "
		sqlStr = sqlStr + " on d.shopid=MS.shopid and d.makerid=MS.makerid "

		sqlStr = sqlStr + " where m.idx<>0"
		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and m.jungsanid='" + FRectJungsanId + "'"
		else
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectMwgubun="M" then
			sqlStr = sqlStr + " and d.chargediv in ('4','8')"
		elseif FRectMwgubun="W" then
			sqlStr = sqlStr + " and d.chargediv in ('2','6')"
		end if

		if FRectonoffgubun="on" then
			sqlStr = sqlStr + " and left(d.shopid,11)='streetshop0'"
		elseif FRectonoffgubun="off" then
			sqlStr = sqlStr + " and left(d.shopid,11)='streetshop8'"
		end if

		sqlStr = sqlStr + " and m.currstate >=0"
		sqlStr = sqlStr + " and m.currstate <8"

		sqlStr = sqlStr + " order by m.yyyymm desc, m.jungsanid, d.shopid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fjungsanid = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				'FItemList(i).Ftotitemcnt = rsget("totitemcnt")
				'FItemList(i).Ftotsum     = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				'FItemList(i).FminusCharge    = rsget("minuscharge")
				'FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				'FItemList(i).Fbigo           = db2html(rsget("bigo"))
				'FItemList(i).Fsegumil        = rsget("segumil")
				'FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FOrgChargeDiv = rsget("orgchargediv")
				FItemList(i).FJungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")

				'FItemList(i).Fgroupidx		= rsget("groupidx")
				'FItemList(i).Ftaxregdate	= rsget("taxregdate")
				'FItemList(i).Fdifferencekey	= rsget("differencekey")
				'FItemList(i).Ftaxtype		= rsget("taxtype")
				'FItemList(i).Ftaxlinkidx	= rsget("taxlinkidx")
				'FItemList(i).Fneotaxno		= rsget("neotaxno")
				FItemList(i).Foffgubun	= rsget("offgubun")

				FItemList(i).FOffMeaip = rsget("offmeaip")
				FItemList(i).FOffChulgo = rsget("offchulgo")
				FItemList(i).FOffSell = rsget("offsell")

				FItemList(i).Fdefaultmargin = rsget("defaultmargin")
				FItemList(i).Fdefaultsuplymargin = rsget("defaultsuplymargin")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetJungsanMasterList
		dim i,sqlStr

		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx<>0"
		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and jungsanid='" + FRectJungsanId + "'"
		else
			sqlStr = sqlStr + " and yyyymm='" + FRectYYYYMM + "'"
		end if
		sqlStr = sqlStr + " and currstate >=0"
		sqlStr = sqlStr + " and currstate <8"
		sqlStr = sqlStr + " and ("
		sqlStr = sqlStr + " 	((offgubun='OFF') and (chargediv in ('0','2','6','8')))"
		sqlStr = sqlStr + " 	or ((offgubun='FRN') and (chargediv in ('0','9')))"
		sqlStr = sqlStr + "	)"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname, p.jungsan_gubun, p.company_no "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"

		sqlStr = sqlStr + " where m.idx<>0"
		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and m.jungsanid='" + FRectJungsanId + "'"
		else
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if
		sqlStr = sqlStr + " and m.currstate >=0"
		sqlStr = sqlStr + " and m.currstate <8"
		sqlStr = sqlStr + " and ("
		sqlStr = sqlStr + " 	((m.offgubun='OFF') and (m.chargediv in ('0','2','6','8')))"
		sqlStr = sqlStr + " 	or ((m.offgubun='FRN') and (m.chargediv in ('0','9')))"
		sqlStr = sqlStr + " )"

		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fjungsanid = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotitemcnt = rsget("totitemcnt")
				FItemList(i).Ftotsum     = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FJungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")

				FItemList(i).Fgroupidx		= rsget("groupidx")
				FItemList(i).Ftaxregdate	= rsget("taxregdate")
				FItemList(i).Fdifferencekey	= rsget("differencekey")
				FItemList(i).Ftaxtype		= rsget("taxtype")
				FItemList(i).Ftaxlinkidx	= rsget("taxlinkidx")
				FItemList(i).Fneotaxno		= rsget("neotaxno")
				FItemList(i).Foffgubun	= rsget("offgubun")

				FItemList(i).Fjungsan_gubun =  rsget("jungsan_gubun")
				FItemList(i).Fcompany_no =  rsget("company_no")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetSubGroupList
		dim i,sqlStr
		sqlStr = "select idx, groupidx, yyyymm, shopid, shopname from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where groupidx=" + CStr(FRectGroupidx)
		sqlStr = sqlStr + " and idx<>" + CStr(FRectIdx)
		sqlStr = sqlStr + " order by idx"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CDetailGroupSubItem

				FItemList(i).Fidx        = rsget("idx")
				FItemList(i).Fgroupidx    = rsget("groupidx")
				FItemList(i).Fyyyymm      = rsget("yyyymm")
				FItemList(i).Fshopid 	 = rsget("shopid")

				FItemList(i).Fshopname   = db2html(rsget("shopname"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub GetOffJungSanDetail()
		dim i,sqlStr

		sqlStr = "select d.*, i.mwdiv from [db_shop].[dbo].tbl_shop_jungsandetail d"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " order by orderno"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					FItemList(i).Fidx            = rsget("idx")
					FItemList(i).Fmasteridx      = rsget("masteridx")
					FItemList(i).Forderno        = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).Fsuplyprice     = rsget("suplyprice")
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fmakerid        = rsget("makerid")
					FItemList(i).Flinkidx        = rsget("linkidx")
					FItemList(i).Fjungsangubun   = rsget("jungsangubun")
					FItemList(i).Fonlinemwdiv   = rsget("mwdiv")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end Sub

	public sub GetOneJungsanMaster()
		dim i,sqlStr
		sqlStr = " select top 1 m.*, p.jungsan_gubun from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and jungsanid='" + FRectJungsanId + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		if Not rsget.Eof then
			set FOneItem = new COffshopJungsanItem

			FOneItem.Fidx = rsget("idx")
			FOneItem.Fyyyymm = rsget("yyyymm")
			FOneItem.Fshopid = rsget("shopid")
			FOneItem.Fjungsanid = rsget("jungsanid")
			FOneItem.Ftotitemcnt = rsget("totitemcnt")
			FOneItem.Ftotsum = rsget("totsum")
			FOneItem.Fminuscharge = rsget("minuscharge")
			FOneItem.Fchargepercent = rsget("chargepercent")
			FOneItem.Frealjungsansum = rsget("realjungsansum")
			FOneItem.Fbigo = db2html(rsget("bigo"))
			FOneItem.Fcurrstate = rsget("currstate")
			FOneItem.Fsegumil = rsget("segumil")
			FOneItem.Fipkumil = rsget("ipkumil")
			FOneItem.Fregdate = rsget("regdate")
			FOneItem.Fchargediv = rsget("chargediv")
			FOneItem.Ffranchargediv = rsget("franchargediv")
			FOneItem.Fgroupidx = rsget("groupidx")
			FOneItem.Ftaxregdate = rsget("taxregdate")
			FOneItem.Fdifferencekey = rsget("differencekey")
			FOneItem.Ftaxtype = rsget("taxtype")
			FOneItem.Ftaxlinkidx = rsget("taxlinkidx")
			FOneItem.Fneotaxno = rsget("neotaxno")
			FOneItem.Foffgubun = rsget("offgubun")

			FOneItem.Fjungsanchargediv = rsget("chargediv")
			FOneItem.Fjungsan_gubun =  rsget("jungsan_gubun")
		end if
		rsget.Close
	end sub

	public Sub GetOffJungSanMakeListByBrand()
		dim i,sqlStr
		dim nextYYYYMMDD

		nextYYYYMMDD = CStr(dateserial(Left(FRectYYYYMM,4),Right(FRectYYYYMM,2)+1,1))

		sqlStr = " select S.userid as shopid, u.makerid, IsNULL(T.totno,0) as totno,"
		sqlStr = sqlStr + " IsNULL(T.totsum,0) as totsum, u.chargediv,u.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user S"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.shopid=S.userid and u.makerid='" + FRectJungsanId + "'"

		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.shopid, "
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectYYYYMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"
		'sqlStr = sqlStr + " 	and year(m.shopregdate)='" + Left(FRectYYYYMM,4) + "'"
		'sqlStr = sqlStr + " 	and month(m.shopregdate)='" + Right(FRectYYYYMM,2) + "'"
		sqlStr = sqlStr + " 	and d.makerid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " 	group by m.shopid"
		sqlStr = sqlStr + " ) as T on S.userid=T.shopid"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " 	on j.jungsanid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " 	and j.shopid=S.userid"
		sqlStr = sqlStr + " 	and j.yyyymm='" + FRectYYYYMM + "'"
		'sqlStr = sqlStr + " 	and (j.yyyymm='2005-12' or (j.yyyymm='2005-11' and j.shopid in ('streetshop871','streetshop872','streetshop873') ) )"

		sqlStr = sqlStr + " where S.userid<>''"

		if FRectOnlyOwnOffJungsan<>"" then
			sqlStr = sqlStr + " and S.shopdiv in ('1','2')"
			'sqlStr = sqlStr + " and S.userid<>'streetshop000'"
		elseif FRectOnlyFranOffJungsan<>"" then
			sqlStr = sqlStr + " and S.shopdiv ='3'"
			'sqlStr = sqlStr + " and S.userid<>'streetshop800'"
		end if
		sqlStr = sqlStr + " order by T.shopid, T.totsum desc, T.totno desc"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffshopJungsanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fjungsanid = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectYYYYMM
					FItemList(i).FSelltotitemcnt = rsget("totno")
					FItemList(i).FSelltotsum     = rsget("totsum")
					FItemList(i).Ftotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Ftotsum     = rsget("jungsantotsum")
					FItemList(i).Fidx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fjungsanchargediv = rsget("chargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")


					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetOldOffJungSanMakeListByBrand()
		dim i,sqlStr
		sqlStr = " select j.shopid, u.makerid, IsNULL(T.totno,0) as totno,"
		sqlStr = sqlStr + " IsNULL(T.totsum,0) as totsum, u.chargediv,u.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster j "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " 	on j.shopid=u.shopid and j.jungsanid=u.makerid and u.makerid='" + FRectJungsanId + "'"

		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.shopid, convert(varchar(7), m.shopregdate,20) as yyyymm,"
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and d.makerid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " 	group by m.shopid, convert(varchar(7), m.shopregdate,20)"
		sqlStr = sqlStr + " ) as T on j.shopid=T.shopid and j.yyyymm=T.yyyymm"

		sqlStr = sqlStr + " where j.jungsanid='" + FRectJungsanId + "'"

		if FRectOnlyOwnOffJungsan<>"" then
			sqlStr = sqlStr + " and j.offgubun='OFF'"
		elseif FRectOnlyFranOffJungsan<>"" then
			sqlStr = sqlStr + " and j.offgubun='FRN'"
		end if
		sqlStr = sqlStr + " order by T.shopid, T.totsum desc, T.totno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffshopJungsanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fjungsanid = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectYYYYMM
					FItemList(i).FSelltotitemcnt = rsget("totno")
					FItemList(i).FSelltotsum     = rsget("totsum")
					FItemList(i).Ftotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Ftotsum     = rsget("jungsantotsum")
					FItemList(i).Fidx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fjungsanchargediv = rsget("chargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")


					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetDesignerJungsan000SumList()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx<>0"

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectYYYYMM<>"" then
			sqlStr = sqlStr + " and yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectOnlyOwnOffJungsan<>"" then
			sqlStr = sqlStr + " and offgubun='OFF'"
		end if

		if FRectOnlyFranOffJungsan<>"" then
			sqlStr = sqlStr + " and offgubun='FRN'"
		end if

		if FRectOnlyUpcheview<>"" then
			sqlStr = sqlStr + " and currstate >0"
			sqlStr = sqlStr + " and chargediv in ('2','6','8','9')"
		end if

		sqlStr = sqlStr + " and currstate <>8"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " where m.idx<>0"

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectYYYYMM<>"" then
			sqlStr = sqlStr + " and yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectOnlyOwnOffJungsan<>"" then
			sqlStr = sqlStr + " and offgubun='OFF'"
			'sqlStr = sqlStr + " and chargediv in ('0','2','4','6','8','9')"
		end if

		if FRectOnlyFranOffJungsan<>"" then
			sqlStr = sqlStr + " and offgubun='FRN'"
			'sqlStr = sqlStr + " and chargediv ='9'"
		end if

		if FRectOnlyUpcheview<>"" then
			sqlStr = sqlStr + " and currstate >0"
			sqlStr = sqlStr + " and chargediv in ('2','6','8','9')"
		end if

		sqlStr = sqlStr + " and currstate <>8"

		sqlStr = sqlStr + " order by m.jungsanid , m.shopid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffshopJungsanItem
				FItemList(i).Fidx = rsget("idx")
				FItemList(i).FShopid	= rsget("shopid")
				FItemList(i).Fjungsanid = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM    = rsget("yyyymm")
				FItemList(i).Ftotitemcnt= rsget("totitemcnt")
				FItemList(i).Ftotsum    = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FjungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")

				FItemList(i).Ftaxtype = rsget("taxtype")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub GetDesignerJungsanListUpcheView()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyOwnOffJungsan="on" then
			sqlStr = sqlStr + " and offgubun='OFF'"
			sqlStr = sqlStr + " and chargediv in ('0','2','6','8')"
		end if

		if FRectOnlyFranOffJungsan="on" then
			sqlStr = sqlStr + " and offgubun='FRN'"
			sqlStr = sqlStr + " and chargediv in ('0','9')"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname, p.jungsan_gubun, p.company_no, p.jungsan_date_off "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"

		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyOwnOffJungsan="on" then
			sqlStr = sqlStr + " and offgubun='OFF'"
			sqlStr = sqlStr + " and chargediv in ('0','2','6','8')"
		end if

		if FRectOnlyFranOffJungsan="on" then
			sqlStr = sqlStr + " and offgubun='FRN'"
			sqlStr = sqlStr + " and chargediv in ('0','9')"
		end if


		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fjungsanid = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotitemcnt = rsget("totitemcnt")
				FItemList(i).Ftotsum     = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FJungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")

				FItemList(i).Fgroupidx		= rsget("groupidx")
				FItemList(i).Ftaxregdate	= rsget("taxregdate")
				FItemList(i).Fdifferencekey	= rsget("differencekey")
				FItemList(i).Ftaxtype		= rsget("taxtype")
				FItemList(i).Ftaxlinkidx	= rsget("taxlinkidx")
				FItemList(i).Fneotaxno		= rsget("neotaxno")
				FItemList(i).Foffgubun	= rsget("offgubun")

				FItemList(i).Fjungsan_gubun =  rsget("jungsan_gubun")
				FItemList(i).Fcompany_no =  rsget("company_no")
				FItemList(i).Fjungsan_date_off =  rsget("jungsan_date_off")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetJungsanListByChargeDiv()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"
		'sqlStr = sqlStr + " and currstate >0"
		'sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectchargediv<>"" then
			sqlStr = sqlStr + " and chargediv='" + FRectchargediv + "'"
		end if

		if FRectCurrState="2" then
			sqlStr = sqlStr + " and currstate<3"
		elseif FRectCurrState<>"" then
			sqlStr = sqlStr + " and currstate='" + FRectCurrState + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname, p.jungsan_gubun, p.company_no "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"

		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"
		'sqlStr = sqlStr + " and currstate >0"
		'sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectchargediv<>"" then
			sqlStr = sqlStr + " and chargediv='" + FRectchargediv + "'"
		end if

		if FRectCurrState="2" then
			sqlStr = sqlStr + " and currstate<3"
		elseif FRectCurrState<>"" then
			sqlStr = sqlStr + " and currstate='" + FRectCurrState + "'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fjungsanid = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotitemcnt = rsget("totitemcnt")
				FItemList(i).Ftotsum     = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FJungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")

				FItemList(i).Fgroupidx		= rsget("groupidx")
				FItemList(i).Ftaxregdate	= rsget("taxregdate")
				FItemList(i).Fdifferencekey	= rsget("differencekey")
				FItemList(i).Ftaxtype		= rsget("taxtype")
				FItemList(i).Ftaxlinkidx	= rsget("taxlinkidx")
				FItemList(i).Fneotaxno		= rsget("neotaxno")
				FItemList(i).Foffgubun	= rsget("offgubun")

				FItemList(i).Fjungsan_gubun =  rsget("jungsan_gubun")
				FItemList(i).Fcompany_no =  rsget("company_no")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 300
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub

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
end class
%>