<%

Class COffShopJaeGo
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Flastrealjeago
	public FIpChulNo
	public FSellNo
	public FJaeGo
	public Ftotipchulno
	public Ftotsellno

	public FMakerID
	public FInputJaeGo

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemID >= 1000000) then
    		getBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopSellByTerm
	public Fidx
	public FItemName
	public FTerm
	public FCount
	public FSum
	public FSpendMile
	public FGainMile

	public FShopid
	public FMakerid
	public FIsBrandShop

	public FJungsanID
	public FSelljungsanID

	public Fsitename
	public Fselltotal
	public Fsellcnt
	public Fdpart
	public Faccountdiv
	public maxt
	public maxc

	public FItemGubun
	public FItemNo
	public FItemID
	public FItemOption
	public FItemCost
	public FItemOptionStr
	public FBuycash
	public FCancelyn
	public FYYYYMMDDHHNNSS

	public FChargeDiv

	public function IsBrandShop()
		''사용안함.
		IsBrandShop = false
		if FIsBrandShop<>"" then IsBrandShop = true
	end function

	public function GetDpartName()
		if Fdpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif Fdpart=2 then
			GetDpartName = "월"
		elseif Fdpart=3 then
			GetDpartName = "화"
		elseif Fdpart=4 then
			GetDpartName = "수"
		elseif Fdpart=5 then
			GetDpartName = "목"
		elseif Fdpart=6 then
			GetDpartName = "금"
		elseif Fdpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Faccountdiv) = "01" then
			JumunMethodName = "현금"
		elseif Cstr(Faccountdiv) = "02" then
			JumunMethodName = "카드"
		end if
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopSellMasterDetailItem
	public ForderNo
	public Ftotalsum
	public Frealsum
	public Fjumunmethod
	public Fshopregdate
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Frealsellprice
	public Fitemno
	public FMakerID
	public Fpointuserno

	Public function JumunMethodColor()
		if Cstr(Fjumunmethod) = "01" then
			JumunMethodColor = "#000000"
		elseif Cstr(Fjumunmethod) = "02" then
			JumunMethodColor = "#0000FF"
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Fjumunmethod) = "01" then
			JumunMethodName = "현금"
		elseif Cstr(Fjumunmethod) = "02" then
			JumunMethodName = "카드"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopSellDetailItem
	public FIdx
	public FShopID
	public FMakerID
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fitemno
	public Fsellprice
	public Frealsellprice
	public Fsubtotal
	public FShopregDate
	public Fsuplyprice
	public FOrderNo

	public Fjungsanid
	public Fcurrentitemprice

	public FOnlineMwDiv

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemID >= 1000000) then
    		getBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class COffShopSellMasterItem
	public Fidx
	public Forderno
	public Fshopid
	public Ftotalsum
	public Frealsum
	public Fjumundiv
	public Fjumunmethod
	public Fshopregdate
	public Fcancelyn
	public Fregdate
	public Fshopidx

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class



class COffShopRealJaegoDetail
	public Fidx
	public Fmasteridx
	public Fmakerid
	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	public Frealjeago
	public Fcancelyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CFranJungSanItem
	public FjungsanMasterIdx

	public Fidx
	public Fbaljuid
	public Fchargeuser
	public FYYYYMM
	public FTotNo
	public FTotalSellcash
	public FTotalBuyCash
	public Fcurrstate

	public FChargeDiv
	public Fjumundivcode
	public FDefaultMargin
	public Fipgodate
	public Fshopid

	public FjungsaMasterIdx

	public Fjungsantotitemcnt
	public Fjungsantotsum
	public FminusCharge
	public FChargePercent
	public FRealjungsansum
	public Fbigo
	public Fsegumil
	public Fipkumil
	public FJungsanChargediv

	public FconvCount
	public FconvSum
	public Flinkidx

	public function GetJumunDivName()
		if Fjumundivcode="101" then
			GetJumunDivName = "가맹점용 개별매입"
		elseif Fjumundivcode="111" then
			GetJumunDivName = "가맹점용 개별특정"
		elseif Fjumundivcode="121" then
			GetJumunDivName = "온라인특정재고->가맹점용특정"
		elseif Fjumundivcode="131" then
			GetJumunDivName = "온라인특정재고->가맹점용매입"
		elseif Fjumundivcode="201" then
			GetJumunDivName = "온라인매입재고->가맹점용매입"
		elseif Fjumundivcode="300" then
			GetJumunDivName = "온라인주문"
		elseif Fjumundivcode="501" then
			GetJumunDivName = "직영샾주문"
		elseif Fjumundivcode="502" then
			GetJumunDivName = "수수료샾"
		elseif Fjumundivcode="503" then
			GetJumunDivName = "프랜차이즈"
		else
			GetJumunDivName = ""
		end if
	end function

	public function GetJumunDivColor()
		if Fjumundivcode="101" then
			GetJumunDivColor = "#0000AA"
		elseif Fjumundivcode="111" then
			GetJumunDivColor = "#AA0000"
		elseif Fjumundivcode="121" then
			GetJumunDivColor = "#AA00AA"
		elseif Fjumundivcode="131" then
			GetJumunDivColor = "#00AAAA"
		elseif Fjumundivcode="201" then
			GetJumunDivColor = "#AAAA00"
		elseif Fjumundivcode="300" then
			GetJumunDivColor = "#FF0000"
		elseif Fjumundivcode="501" then
			GetJumunDivColor = "#0000FF"
		elseif Fjumundivcode="502" then
			GetJumunDivColor = "#00FF00"
		elseif Fjumundivcode="503" then
			GetJumunDivColor = "#AAFFAA"
		else
			GetJumunDivColor = "#000000"
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
		elseif Fcurrstate="9" then
			GetStateColor = "#888844"
		elseif Fcurrstate=" " then
			GetStateColor = "#AAAAAA"
		else

		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "10x10 특정"
		elseif FChargeDiv="4" then
			getChargeDivName = "10x10 매입"
		elseif FChargeDiv="5" then
			getChargeDivName = "출고분정산"
		elseif FChargeDiv="6" then
			getChargeDivName = "업체 특정"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체 매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 특정"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		else
			getJungSanChargeDivName = FJungsanChargediv
		end if
	end function

	public function getJungSanChargeDivNameUpcheView()
		if FJungsanChargediv="2" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivNameUpcheView = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivNameUpcheView = "통합"
		else
			getJungSanChargeDivNameUpcheView = FJungsanChargediv
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class COffShopJungSanItem
	public Fidx
	public FShopid

	public Fchargeuser
	public Fchargename
	public FYYYYMM
	public Ftotno
	public FSum

	public FjungsaMasterIdx
	public Fcurrstate

	public Fjungsantotitemcnt
	public Fjungsantotsum
	public FminusCharge
	public FChargePercent
	public FRealjungsansum
	public Fbigo
	public Fsegumil
	public Fipkumil

	public Fchargediv
	public Fjungsan_acctname
	public Fjungsan_bank
	public Fjungsan_acctno
	public Fcompany_name
	public FJungsanChargediv
	public Fjungsan_date_off
	public Fjungsan_date_frn

	public FNoFixSum
	public FFixSum
	public FIpkumSum
	public FAutoJungsan

	public FShopName
	public FShopDiv
	public Fdefaultmargin

	public FFranChargeDiv

	public FGroupidx
	public FTaxRegdate
	public FDifferencekey
	public FTaxType
	public FTaxLinkidx
	public Fneotaxno
	public Foffgubun

	public Fonlinedefaultmargine
	public Fonlinemaeipdiv

	public Ftotchulgono
	public FtotchulgoSum

	public function GetOnlineMaeipDivName
		if Fonlinemaeipdiv="M" then
			GetOnlineMaeipDivName = "매입"
		elseif Fonlinemaeipdiv="W" then
			GetOnlineMaeipDivName = "특정"
		elseif Fonlinemaeipdiv="U" then
			GetOnlineMaeipDivName = "업체"
		else
			GetOnlineMaeipDivName = Fonlinemaeipdiv
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
			GetStateColor = "#AAAAAA"
		else

		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "10x10 특정"
		elseif FChargeDiv="4" then
			getChargeDivName = "10x10 매입"
		elseif FChargeDiv="5" then
			getChargeDivName = "출고분정산"
		elseif FChargeDiv="6" then
			getChargeDivName = "업체 특정"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체 매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		else
			getJungSanChargeDivName = FJungsanChargediv
		end if
	end function

	public function getJungSanChargeDivNameUpcheView()
		if FJungsanChargediv="2" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivNameUpcheView = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivNameUpcheView = "통합"
		else
			getJungSanChargeDivNameUpcheView = FJungsanChargediv
		end if
	end function

	public function GetFranChargeDivName()
		if FFranChargeDiv="2" then
			GetFranChargeDivName = "특정"
		elseif FFranChargeDiv="4" then
			GetFranChargeDivName = "매입"
		elseif FFranChargeDiv="5" then
			GetFranChargeDivName = "매입"
		elseif FFranChargeDiv="6" then
			GetFranChargeDivName = "특정"
		elseif FFranChargeDiv="8" then
			GetFranChargeDivName = "매입"
		else
			GetFranChargeDivName = FFranChargeDiv
		end if
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

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

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

	public function getDetailGubunName()
		if Fjungsangubun = "101" then
			getDetailGubunName = "매입"
		elseif Fjungsangubun = "131" then
			getDetailGubunName = "특정재고->매입"
		elseif Fjungsangubun = "111" then
			getDetailGubunName = "특정"
		elseif Fjungsangubun = "121" then
			getDetailGubunName = "특정"
		elseif Fjungsangubun = "801" then
			getDetailGubunName = "off매입"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffshopAutoJungsan
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectYYYYMM
	public FRectMakerid
	public FRectJungsanYYYY
	public FRectJungsanMM
	public FRectShopID

	public FRectOnlyMaeipChecked

	public sub GetFranChulgoJungsanList
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))


		sqlStr = sqlStr + " select d.makerid, d.shopid,"
		sqlStr = sqlStr + " IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and j.jungsanid=d.makerid"
			sqlStr = sqlStr + " and j.shopid=d.shopid"

			sqlStr = sqlStr + " left join ("
			sqlStr = sqlStr + " select m.socid as shopid, count(itemno) as totno, sum(buycash*itemno*-1) as totsum "
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
			sqlStr = sqlStr + " where m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
			sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " and m.ipchulflag='S'"
			sqlStr = sqlStr + " and m.code=d.mastercode"
			sqlStr = sqlStr + " and m.deldt is null"
			sqlStr = sqlStr + " and Left(m.socid,11)='streetshop8'"
			sqlStr = sqlStr + " and d.imakerid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and d.mwgubun='C'"
			sqlStr = sqlStr + " and d.deldt is null"
			sqlStr = sqlStr + " group by m.socid"

			sqlStr = sqlStr + " ) T on T.shopid=d.shopid"
		sqlStr = sqlStr + " where Left(d.shopid,11) ='streetshop8'"
		sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and (T.totno<>0 or j.totitemcnt<>0)"
		sqlStr = sqlStr + " and d.chargediv in ('4','5')"
		sqlStr = sqlStr + " order by d.makerid, d.shopid"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem

					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totsum")
					'FItemList(i).FTotalBuyCash  = 0
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					'FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fdefaultmargin	= rsget("defaultmargin")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fshopid = rsget("shopid")


					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public sub GetFranMeaipTargetListConv
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select m.id, d.imakerid, sum(d.itemno*-1) as ccnt,"
		sqlStr = sqlStr + " sum(d.itemno*d.sellcash*-1) as totalsellcash,"
		sqlStr = sqlStr + " sum(d.itemno*d.buycash*-1) as totalbuycash, m.executedt"
		sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		if FRectOnlyMaeipChecked="on" then
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.iitemgubun='10' and d.itemid=i.itemid"
		end if
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and d.imakerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and Left(m.code,2)='SO'"
		sqlStr = sqlStr + " and Left(m.socid,11)='streetshop8'"

		if FRectOnlyMaeipChecked="on" then
			sqlStr = sqlStr + " and i.mwdiv='W'"
		end if

		sqlStr = sqlStr + " group by m.id, d.imakerid,m.executedt"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("id")
					''FItemList(i).Flinkidx = rsget("linkidx")
					'FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("imakerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("executedt")


					'FItemList(i).Fshopid = rsget("shopid")
					'FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).FYYYYMM  = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).Ftotno		= rsget("totno")
					'FItemList(i).FTotalSellcash = rsget("totalsellcash")
					'FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					'FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					'FItemList(i).Fjumundivcode = rsget("jdivcode")
					'FItemList(i).FChargeDiv    = rsget("chargediv")
					'FItemList(i).Fcurrstate    = rsget("currstate")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipTargetListByIpgo()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select m.socid, m.id , m.totalsellcash"
		sqlStr = sqlStr + " ,m.totalbuycash , m.executedt, m.divcode as jdivcode "
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		'sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		'sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash"
		sqlStr = sqlStr + " ,j.currstate"
		sqlStr = sqlStr + " ,d.chargediv,d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.socid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on m.socid=d.makerid"
			sqlStr = sqlStr + " and d.shopid='streetshop800'"
		sqlStr = sqlStr + " where socid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and executedt<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(executedt)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(executedt)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and divcode ='801'"
		sqlStr = sqlStr + " and deldt is NULL"
		'sqlStr = sqlStr + " group by targetid, j.idx, j.totsum, j.realjungsansum, j.currstate, j.chargediv"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("id")
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("socid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("executedt")


					'FItemList(i).Fshopid = rsget("shopid")
					'FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).FYYYYMM  = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).Ftotno		= rsget("totno")
					'FItemList(i).FTotalSellcash = rsget("totalsellcash")
					'FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fcurrstate    = rsget("currstate")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetFranMeaipTargetList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select m.targetid, m.idx , m.totalsellcash"
		sqlStr = sqlStr + " ,m.totalbuycash , m.ipgodate, m.divcode as jdivcode "
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		'sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		'sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash"
		sqlStr = sqlStr + " ,j.currstate"
		sqlStr = sqlStr + " ,d.chargediv,d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.targetid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on m.targetid=d.makerid"
			sqlStr = sqlStr + " and d.shopid='streetshop800'"
		sqlStr = sqlStr + " where baljuid='10x10'"
		sqlStr = sqlStr + " and targetid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and statecd='9'"
		sqlStr = sqlStr + " and ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and ipgodate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(ipgodate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(ipgodate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and divcode in ('101','131')"
		sqlStr = sqlStr + " and deldt is NULL"
		'sqlStr = sqlStr + " group by targetid, j.idx, j.totsum, j.realjungsansum, j.currstate, j.chargediv"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("targetid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("ipgodate")


					'FItemList(i).Fshopid = rsget("shopid")
					'FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).FYYYYMM  = FRectJungsanYYYY + "-" + FRectJungsanMM
					'FItemList(i).Ftotno		= rsget("totno")
					'FItemList(i).FTotalSellcash = rsget("totalsellcash")
					'FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fcurrstate    = rsget("currstate")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetFranWitakTargetList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select sd.shopid, sd.makerid "
		sqlStr = sqlStr + " ,count(d.itemno) as totno"
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash, j.currstate,j.chargediv as jdivcode"
		sqlStr = sqlStr + " ,sd.chargediv, sd.defaultmargin"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_designer sd"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=sd.makerid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and Left(j.shopid,11)='streetshop8'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " where m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(m.shopregdate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(m.shopregdate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopid =sd.shopid"
		sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop8'"
		sqlStr = sqlStr + " and d.makerid=sd.makerid"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and sd.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and sd.chargediv in ('2','6')"

		sqlStr = sqlStr + " group by sd.shopid, sd.makerid,j.idx,j.totsum,j.realjungsansum,j.currstate,j.chargediv"
		sqlStr = sqlStr + " ,sd.chargediv, sd.defaultmargin"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					'FItemList(i).Fidx = rsget("idx")
					FItemList(i).Fshopid = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM  = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno		= rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					'FItemList(i).Fipgodate	   = rsget("ipgodate")
					FItemList(i).Fcurrstate    = rsget("currstate")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipTargetWitakList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid, d.makerid, "
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=m.targetid"
		sqlStr = sqlStr + " and j.shopid='streetshop800'"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and u.makerid=m.targetid and u.shopid='streetshop800' "

		sqlStr = sqlStr + " where m.baljuid='10x10'"
		sqlStr = sqlStr + " and m.targetid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and m.statecd='9'"
		sqlStr = sqlStr + " and m.ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.ipgodate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(m.ipgodate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(m.ipgodate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.divcode in ('101','131')"
		sqlStr = sqlStr + " and m.deldt is NULL"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fidx = rsget("idx")
					'FItemList(i).Fbaljuid = rsget("baljuid")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fipgodate	   = rsget("ipgodate")
					FItemList(i).Fcurrstate    = rsget("currstate")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close


	end sub

	public Sub GetTargetList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select T.*, u.chargediv,u.autojungsan,u.defaultmargin,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " s.shopname, s.shopdiv"

		sqlStr = sqlStr + " from ("

		sqlStr = sqlStr + " select m.shopid, d.makerid, "
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(m.shopregdate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(m.shopregdate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		'sqlStr = sqlStr + " and m.shopid not in ('cafe001','cafe002','cafe003')"
		sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=T.makerid"
		sqlStr = sqlStr + " and j.shopid=T.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid='" + FRectMakerid + "' and u.makerid=T.makerid and u.shopid=T.shopid "

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s"
		sqlStr = sqlStr + " on s.userid=T.shopid "
		sqlStr = sqlStr + " where s.shopdiv<>'3'"
		''sqlStr = sqlStr + " and u.chargediv<>'4'"
		sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"

''response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")

					FItemList(i).FShopname	 = db2html(rsget("shopname"))
					FItemList(i).FShopDiv	 = rsget("shopdiv")
					FItemList(i).Fdefaultmargin = rsget("defaultmargin")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
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

class COffShopSellReport
	public FItemList()
	public FCountList()
	public FOneJeaGoMaster
	public FOneJungSanMaster

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectStartDay
	public FRectEndDay
	public FRectNormalOnly
	public FRectJungsanId
	public FRectDesigner
	public FRectShopID
	public FRectTerms
	public FRectItemId

	public FRectJungsanYYYYMM
	public FRectJungsanYYYY
	public FRectJungsanMM
	public FYYYYMMDDHHNNSS

	public FRectJaegoNo
	public FRectIDX
	public maxt
	public maxt2
	public maxc
	public FRectPointYN

	public FRectOnlymijungsan
	public FRectOnlyUpcheJungSan
	public FRectOnlyFranUpcheJungSan
	public FRectNotIncludeWonChon
	public FRectOnlyIncludeWonChon
	public FRectOnlyIncludeNoTax

	public FRectOnlyShop
	public FRectNotChargeDiv
	public FRectChargeDiv

	public FRectOffgubun

	public FDayTsellsum
	public FDayTea
	public FRectOrder
	public FRectMWgubun

	public FRectnomeachul

	public FRectUpcheWitakOnly
	public FRectJungsanDate
	public FRectOldData

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function




	public Sub GetNotExistsInSertJungSanMaster()
		dim i,sqlStr, masterExists
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select j.*, u.chargename from [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_chargeuser u"
		sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
		sqlStr = sqlStr + " where j.yyyymm='" + FRectJungsanYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and j.jungsanid='" + FRectJungsanID + "'"
		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masterExists = true
				redim  FItemList(0)

				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FjungsaMasterIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fchargeuser = rsget("jungsanid")
				FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate  = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")

			else
				masterExists = false
			end if
		rsget.Close


		if Not masterExists then
			sqlStr = " insert into [db_shop].[dbo].tbl_shop_jungsanmaster"
			sqlStr = sqlStr + " (yyyymm,shopid,jungsanid,totitemcnt,totsum,currstate)"
			sqlStr = sqlStr + " select '" + FRectJungsanYYYYMM + "','" + FRectShopID +"','" + FRectJungsanID + "',"
			sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum, '0'"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
			sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
			'sqlStr = sqlStr + " and year(m.shopregdate)='" + FRectJungsanYYYY + "'"
			'sqlStr = sqlStr + " and month(m.shopregdate)='" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			rsget.Open sqlStr,dbget,1

			sqlStr = " select j.*, u.chargename from [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_chargeuser u"
			sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
			sqlStr = sqlStr + " where j.yyyymm='" + FRectJungsanYYYYMM + "'"
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectJungsanID + "'"
			rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					masterExists = true
					redim  FItemList(0)

					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FjungsaMasterIdx = rsget("idx")
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("jungsanid")
					FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totitemcnt")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fcurrstate = rsget("currstate")

					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FChargePercent  = rsget("chargepercent")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fbigo           = db2html(rsget("bigo"))
					FItemList(i).Fsegumil        = rsget("segumil")
					FItemList(i).Fipkumil        = rsget("ipkumil")
				else
					masterExists = false
				end if
			rsget.Close
		end if
	end Sub

	public Sub GetDesignerJungsanList()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv in ('2','6','8')"
		end if

		if FRectOnlyFranUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv ='9'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"

		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv in ('2','6','8')"
		end if

		if FRectOnlyFranUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv ='9'"
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
				FItemList(i).FjungsaMasterIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				'FItemList(i).Fchargeuser = rsget("jungsanid")
				'FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
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
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub GetMiJungSanList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid, u.chargeuser, u.chargename,"
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,IsNull(j.minuscharge,0) as minuscharge,IsNull(j.realjungsansum,0) as realjungsansum"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_chargeuser u"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
		'sqlStr = sqlStr + " and year(m.shopregdate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(m.shopregdate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and d.jungsanid=u.chargeuser"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " group by m.shopid,u.chargeuser,u.chargename,j.idx,j.currstate,j.minuscharge,j.realjungsansum"
		sqlStr = sqlStr + " order by totsum desc,totno desc"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("chargeuser")
					FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetJungsanSummaryList()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, m.shopid, sum(IsNull(m.totitemcnt,0)) as totitemcnt,"
		sqlStr = sqlStr + " sum(IsNull(m.totsum,0)) as totsum,"
		sqlStr = sqlStr + " sum(IsNull(m.minuscharge,0)) as minuscharge, "
		sqlStr = sqlStr + " sum(IsNull(m.chargepercent,0)) as chargepercent, "
		sqlStr = sqlStr + " sum(IsNull(m.realjungsansum,0)) as realjungsansum, m.currstate"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		sqlStr = sqlStr + " where m.currstate<7"
		sqlStr = sqlStr + " group by m.yyyymm, m.shopid, m.currstate"
		sqlStr = sqlStr + " order by m.yyyymm desc, m.currstate"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				'FItemList(0).Fbigo           = db2html(rsget("bigo"))
				'FItemList(0).Fsegumil        = rsget("segumil")
				'FItemList(0).Fipkumil        = rsget("ipkumil")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetJungsanSummaryListByChargeDiv()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, m.chargediv,  "
		sqlStr = sqlStr + " Sum (case when m.currstate='7' then m.realjungsansum"
		sqlStr = sqlStr + " else 0 end ) as ipkumsum, "
		sqlStr = sqlStr + " Sum (case when (m.currstate='3') then m.realjungsansum"
		sqlStr = sqlStr + " else 0 end ) as fixsum, "
		sqlStr = sqlStr + " Sum (case when (m.currstate='0') or (m.currstate='1') or (m.currstate='2') then m.realjungsansum"
		''sqlStr = sqlStr + " when (m.currstate=' ') then m.totsum"
		sqlStr = sqlStr + " else 0 end ) as nofixsum "

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " where m.regdate>'2004-01-01'"
		sqlStr = sqlStr + " group by m.yyyymm, m.chargediv"
		sqlStr = sqlStr + " order by m.yyyymm desc, m.chargediv"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Fchargediv = rsget("chargediv")

				FItemList(i).FNoFixSum  = rsget("nofixsum")
				FItemList(i).FFixSum    = rsget("fixsum")
				FItemList(i).FIpkumSum  = rsget("ipkumsum")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetJungsanFix26MasterList()
		dim i,sqlStr
		sqlStr = " select m.*, p.jungsan_acctname, p.jungsan_bank, p.jungsan_acctno, p.company_name, p.jungsan_date_off, p.jungsan_date_frn"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		if FRectJungsanDate="w" then
			sqlStr = sqlStr + " left join (select distinct makerid "
			sqlStr = sqlStr + " 			from [db_shop].[dbo].tbl_shop_designer "
			sqlStr = sqlStr + " 			where chargediv='6' "
			sqlStr = sqlStr + " ) as U on m.jungsanid=U.makerid"
		end if

		sqlStr = sqlStr + " where currstate='3'"

		sqlStr = sqlStr + " and m.chargediv in ('2','6','8','0','9')"

		if FRectOffgubun<>"" then
			sqlStr = sqlStr + " and m.offgubun='" + FRectOffgubun + "'"
		end if

		'sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"

		if (FRectJungsanDate="n") then
			sqlStr = sqlStr + " and (((Left(m.shopid,11)='streetshop0') and (IsNULL(p.jungsan_date_off,'')='')) or  ((Left(m.shopid,11)='streetshop8') and (IsNULL(p.jungsan_date_frn,'')='')))"
		elseif ((FRectJungsanDate="15일") or (FRectJungsanDate="말일")) then
			sqlStr = sqlStr + " and (((Left(m.shopid,11)='streetshop0') and (p.jungsan_date_off='" + FRectJungsanDate + "')) or  ((Left(m.shopid,11)='streetshop8') and (p.jungsan_date_frn='" + FRectJungsanDate + "')))"
		elseif FRectJungsanDate<>"" then
			sqlStr = sqlStr + " and U.makerid is not null"
		end if

		sqlStr = sqlStr + " order by m.yyyymm asc, m.taxregdate, m.jungsanid "

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).Fidx	 = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FChargeUser = rsget("jungsanid")

				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")


				FItemList(i).Fjungsan_acctname = rsget("jungsan_acctname")
				FItemList(i).Fjungsan_bank = rsget("jungsan_bank")
				FItemList(i).Fjungsan_acctno = rsget("jungsan_acctno")
				FItemList(i).Fcompany_name = rsget("company_name")
				FItemList(i).Fjungsan_date_off = rsget("jungsan_date_off")
				FItemList(i).Fjungsan_date_frn = rsget("jungsan_date_frn")


				FItemList(i).FGroupidx      = rsget("groupidx")
				FItemList(i).FTaxRegdate    = rsget("taxregdate")
				FItemList(i).FDifferencekey = rsget("differencekey")
				FItemList(i).FTaxType       = rsget("taxtype")
				FItemList(i).FTaxLinkidx    = rsget("taxlinkidx")
				FItemList(i).Fneotaxno      = rsget("neotaxno")
				FItemList(i).Foffgubun      = rsget("offgubun")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetJungsanFixMasterList()
		dim i,sqlStr
		sqlStr = " select m.*, p.jungsan_acctname, p.jungsan_bank, p.jungsan_acctno,p.company_name "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		sqlStr = sqlStr + " where currstate='3'"

		if FRectNotChargeDiv<>"" then
			sqlStr = sqlStr + " and m.chargediv='" + FRectNotChargeDiv + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'면세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		sqlStr = sqlStr + " order by m.idx desc"
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).Fidx	 = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FChargeUser = rsget("jungsanid")

				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")


				FItemList(i).Fjungsan_acctname = rsget("jungsan_acctname")
				FItemList(i).Fjungsan_bank = rsget("jungsan_bank")
				FItemList(i).Fjungsan_acctno = rsget("jungsan_acctno")
				FItemList(i).Fcompany_name = rsget("company_name")

				FItemList(i).FGroupidx      = rsget("groupidx")
				FItemList(i).FTaxRegdate    = rsget("taxregdate")
				FItemList(i).FDifferencekey = rsget("differencekey")
				FItemList(i).FTaxType       = rsget("taxtype")
				FItemList(i).FTaxLinkidx    = rsget("taxlinkidx")
				FItemList(i).Fneotaxno      = rsget("neotaxno")
				FItemList(i).Foffgubun      = rsget("offgubun")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetOneJungsanMaster()
		dim i,sqlStr
		sqlStr = " select * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + FRectIdx + ""
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
				set FItemList(0) = new COffShopJungSanItem
				FItemList(0).FjungsaMasterIdx = rsget("idx")
				FItemList(0).FShopid	 = rsget("shopid")
				'FItemList(0).Fchargeuser = rsget("jungsanid")
				'FItemList(0).Fchargename = rsget("chargename")
				FItemList(0).FYYYYMM     = rsget("yyyymm")
				FItemList(0).Ftotno      = rsget("totitemcnt")
				FItemList(0).FSum        = rsget("totsum")
				FItemList(0).Fcurrstate = rsget("currstate")

				FItemList(0).FminusCharge    = rsget("minuscharge")
				FItemList(0).FChargePercent  = rsget("chargepercent")
				FItemList(0).FRealjungsansum = rsget("realjungsansum")
				FItemList(0).Fbigo           = db2html(rsget("bigo"))
				FItemList(0).Fsegumil        = rsget("segumil")
				FItemList(0).Fipkumil        = rsget("ipkumil")
		end if

		rsget.Close
	end Sub

	public Sub GetOffJungSanDetailSum()
		dim i,sqlStr
		sqlStr = "select itemgubun,itemid,itemoption,itemname,itemoptionname,sellprice,realsellprice,suplyprice,sum(itemno) as itemno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsandetail"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " group by itemgubun,itemid,itemoption,itemname,itemoptionname,sellprice,realsellprice,suplyprice"
		sqlStr = sqlStr + " order by itemgubun,itemid,itemoption"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).Fsuplyprice     = rsget("suplyprice")
					FItemList(i).Fitemno         = rsget("itemno")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end sub

	public Sub GetOffJungSanDetail()
		dim i,sqlStr
		dim isEof

		sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		if FRectJungsanId<>"" then
				sqlStr = sqlStr + " and jungsanid='" + CStr(FRectJungsanId) + "'"
		end if

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			set FOneJungSanMaster = new COffShopJungsanItem

			FOneJungSanMaster.FShopid	 = rsget("shopid")
			FOneJungSanMaster.Fchargeuser = rsget("jungsanid")
			FOneJungSanMaster.FYYYYMM     = rsget("yyyymm")
			FOneJungSanMaster.Ftotno      = rsget("totitemcnt")
			FOneJungSanMaster.FSum        = rsget("totsum")
			FOneJungSanMaster.Fcurrstate = rsget("currstate")
			FOneJungSanMaster.FminusCharge    = rsget("minuscharge")
			FOneJungSanMaster.FRealjungsansum = rsget("realjungsansum")

			FOneJungSanMaster.Fchargediv = rsget("chargediv")

			FOneJungSanMaster.FSegumil = rsget("segumil")
			FOneJungSanMaster.Fipkumil = rsget("ipkumil")
			'FOneJungSanMaster.Fregdate = rsget("regdate")
		else
			isEof = true
		end if

		rsget.Close

		if isEof then dbget.close()	:	response.End

		sqlStr = "select * from [db_shop].[dbo].tbl_shop_jungsandetail"
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

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end Sub


	'' 가맹점 특정->매입 출고분 정산내역 검토
	public sub GetFranWitak2MeaipChulgoJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select C.shopid , C.makerid, C.totcnt ,"
		sqlStr = sqlStr + " C.totalsellcash, "
		sqlStr = sqlStr + " C.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(j.idx,'') as jungsanmasteridx,"
		sqlStr = sqlStr + " IsNULL(j.currstate,'') as currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " IsNull(j.chargediv,'') as jchargediv," 			'' (정산당시 정산 구분)
		sqlStr = sqlStr + " IsNull(d.chargediv,'') as chargediv, "			'' (현재 정산 구분)
		sqlStr = sqlStr + " IsNull(d.defaultmargin,0) as defaultmargin"		'' (현재 마진)
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ( " '' 가맹점 특정 -> 매입 출고 내역
		sqlStr = sqlStr + " 	select m.socid as shopid, d.imakerid as makerid, sum(d.itemno*-1) as totcnt, sum(d.itemno*d.sellcash*-1) as totalsellcash, sum(d.itemno*d.buycash*-1) as totalbuycash from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.ipchulflag='S'"
		sqlStr = sqlStr + " 	and Left(m.socid,11)='streetshop8'"
		sqlStr = sqlStr + " 	and d.mwgubun='C'"
		sqlStr = sqlStr + " 	group by m.socid , d.imakerid"
		sqlStr = sqlStr + " ) as C"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " 	on C.shopid=d.shopid and C.makerid=d.makerid"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " 	on C.shopid=j.shopid"
		sqlStr = sqlStr + " 	and C.makerid=j.jungsanid"
		sqlStr = sqlStr + " 	and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " where d.chargediv<>'2'"
		sqlStr = sqlStr + " and d.chargediv<>'6'"
		sqlStr = sqlStr + " order by C.makerid, C.shopid "

''response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).Fshopid = rsget("shopid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fdefaultmargin = rsget("defaultmargin")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipJungSanAutoList2()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select C.shopid , C.makerid, C.totcnt ,"
		sqlStr = sqlStr + " C.totalsellcash, "
		sqlStr = sqlStr + " C.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(j.idx,'') as jungsanmasteridx,"
		sqlStr = sqlStr + " IsNULL(j.currstate,'') as currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " IsNull(j.chargediv,'') as jchargediv," 			'' (정산당시 정산 구분)
		sqlStr = sqlStr + " IsNull(d.chargediv,'') as chargediv, "			'' (현재 정산 구분)
		sqlStr = sqlStr + " IsNull(d.defaultmargin,0) as defaultmargin"		'' (현재 마진)
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ( " '' 가맹점 특정 -> 매입 출고 내역
		sqlStr = sqlStr + " 	select m.socid as shopid, d.imakerid as makerid, sum(d.itemno) as totcnt, sum(d.itemno*d.sellcash) as totalsellcash, sum(d.itemno*d.buycash) as totalbuycash from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.ipchulflag='I'"
		sqlStr = sqlStr + " 	and m.divcode='801'"
		sqlStr = sqlStr + " 	group by m.socid , d.imakerid"
		sqlStr = sqlStr + " ) as C"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " 	on C.shopid=d.shopid and C.makerid=d.makerid"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " 	on j.shopid='streetshop800'"
		sqlStr = sqlStr + " 	and C.makerid=j.jungsanid"
		sqlStr = sqlStr + " 	and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " order by C.makerid, C.shopid "


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).Fshopid = "streetshop800"
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")

					FItemList(i).Fdefaultmargin = rsget("defaultmargin")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end sub

	public Sub GetFranMeaipJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select  targetid,count(m.idx) as totcnt,"
		sqlStr = sqlStr + " sum(m.totalsellcash) as totalsellcash, sum(m.totalbuycash) as totalbuycash,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.targetid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"

			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on d.shopid='streetshop800'"
			sqlStr = sqlStr + " and d.makerid=m.targetid"

		sqlStr = sqlStr + " where m.baljuid='10x10'"
		sqlStr = sqlStr + " and m.statecd='9'"

		sqlStr = sqlStr + " and m.ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.ipgodate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and year(m.ipgodate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and month(m.ipgodate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.divcode in ('101','131')"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " group by m.targetid, j.idx, j.currstate, j.totitemcnt"
		sqlStr = sqlStr + " ,j.totsum, j.minuscharge, j.realjungsansum, j.chargediv"
		sqlStr = sqlStr + " ,d.chargediv, d.defaultmargin"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("targetid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")

					FItemList(i).Fdefaultmargin = rsget("defaultmargin")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetOffJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select T.*, u.chargediv,u.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " j.totitemcnt as jungsantotitemcnt,"
		sqlStr = sqlStr + " j.totsum as jungsantotsum,"
		sqlStr = sqlStr + " IsNULL(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv"
		sqlStr = sqlStr + " from ("

		sqlStr = sqlStr + " 	select m.shopid, d.makerid, "
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=T.makerid"
		sqlStr = sqlStr + " and j.shopid=T.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid=T.makerid and u.shopid=T.shopid "

		sqlStr = sqlStr + " where j.idx is null"

		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and u.chargediv='" + FRectChargeDiv + "'"
		end if

		sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetFranWitakJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select d.makerid, d.shopid,"
		sqlStr = sqlStr + " IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=d.makerid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and Left(j.shopid,11) ='streetshop8'"
			sqlStr = sqlStr + " and j.shopid=d.shopid"

			sqlStr = sqlStr + " left join ("
			sqlStr = sqlStr + " select d.makerid, m.shopid, count(itemno) as totno, sum(sellprice*itemno) as totsum "
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
			sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"

			'sqlStr = sqlStr + " and year(m.shopregdate)='" + FRectJungsanYYYY + "'"
			'sqlStr = sqlStr + " and month(m.shopregdate)='" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and Left(m.shopid,11) ='streetshop8'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			end if
			sqlStr = sqlStr + " group by d.makerid, m.shopid"
			sqlStr = sqlStr + " ) T on T.makerid=d.makerid and T.shopid=d.shopid"
		sqlStr = sqlStr + " where Left(d.shopid,11) ='streetshop8'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if

	    sqlStr = sqlStr + " and (T.totno<>0 or j.totitemcnt<>0)"
		sqlStr = sqlStr + " and d.chargediv in ('2','6')"
		sqlStr = sqlStr + " order by d.makerid, d.shopid"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totsum")
					'FItemList(i).FTotalBuyCash  = 0
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					'FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fdefaultmargin	= rsget("defaultmargin")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fshopid = rsget("shopid")



					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetOffJungSanList()
		dim i,sqlStr
		dim nextYYYYMMDD

		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select S.shopid, S.makerid, IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " IsNULL(T2.totno,0) as totchulgono, IsNULL(T2.totsum,0) as totchulgosum,"
		sqlStr = sqlStr + " S.chargediv,S.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv, u.defaultmargine, u.maeipdiv "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer S"

		'' 판매내역
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.shopid, d.makerid, "
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"

		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T on S.shopid=T.shopid and S.makerid=T.makerid"

		'' 출고내역
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select  m.socid, d.imakerid,"
		sqlStr = sqlStr + " 	count(d.itemno*-1) as totno, sum(d.buycash*d.itemno*-1) as totsum"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.ipchulflag='S'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " 	and m.socid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " 	and d.imakerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	group by m.socid, d.imakerid"
		sqlStr = sqlStr + " ) as T2 on S.shopid=T2.socid and S.makerid=T2.imakerid"

		'' 정산내역
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=S.makerid"
		sqlStr = sqlStr + " and j.shopid=S.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and j.jungsanid='" + FRectDesigner + "'"
		end if

		'' 온라인마진
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c u"
		sqlStr = sqlStr + " on S.makerid=u.userid "


		'' 쿼리조건
		sqlStr = sqlStr + " where 1=1"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and S.shopid='" + FRectShopID + "'"
		end if
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and S.makerid='" + FRectDesigner + "'"
		end if
		if FRectOnlymijungsan="on" then
			sqlStr = sqlStr + " and j.currstate is NULL"
		end if
		if FRectnomeachul="on" then
			sqlStr = sqlStr + " and IsNULL(T.totsum,0)<>0"
		end if

		sqlStr = sqlStr + " order by S.shopid, S.chargediv desc, T.totsum desc, T.totno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")


					FItemList(i).Ftotchulgono     = rsget("totchulgono")
					FItemList(i).FtotchulgoSum    = rsget("totchulgosum")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")

					FItemList(i).Fonlinedefaultmargine 	= rsget("defaultmargine")
					FItemList(i).Fonlinemaeipdiv		= rsget("maeipdiv")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetCurrentJaeGoMinusList()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno,"
		sqlStr = sqlStr + " i.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
				sqlStr = sqlStr + " and m.chargeid='" + FRectDesigner + "'"
			end if

			if FRectShopID<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			end if
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.jungsan='" + FRectDesigner + "'"
			end if
			if FRectShopID<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			end if
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"
		sqlStr = sqlStr + " where i.shopitemid<>0"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"
		sqlStr = sqlStr + " and ((IsNull(S.itemno,0)-IsNull(T.itemno,0))<" + CStr(FRectJaegoNo) + ")"
		sqlStr = sqlStr + " order by i.makerid"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					FItemList(i).FMakerID		  = rsget("makerid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetCurrentJaeGoList1()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"

		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close



	end Sub

	public Sub GetCurrentJaeGoList2()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
			'sqlStr = sqlStr + " and m.chargeid='" + FRectDesigner + "'"
			if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"

		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetBrandMaeipItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid,m.idx as orderno, d.idx, d.sellcash as sellprice, d.itemno,d.sellcash as realsellprice, d.designerid as makerid,"
		sqlStr = sqlStr + " m.chargeid as jungsanid, d.itemgubun,d.shopitemid as itemid,d.itemoption, i.shopitemname as itemname,i.shopitemoptionname as itemoptionname, "
		sqlStr = sqlStr + " m.execdt as shopregdate, d.suplycash as suplyprice, oi.mwdiv as onlinemwdiv"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shop_ipchul_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " on d.itemgubun=i.itemgubun and d.shopitemid=i.shopitemid and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.itemgubun='10' and d.shopitemid=oi.itemid"

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.execdt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.execdt<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and Year(m.execdt)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and Month(m.execdt)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopId + "'"
		sqlStr = sqlStr + " and m.chargeid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " and m.statecd>=7"
		sqlStr = sqlStr + " and m.deleteyn='N'"
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " order by d.idx"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).FIdx			 = rsget("idx")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")

					FItemList(i).FOnlineMwDiv = rsget("onlinemwdiv")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub


	public Sub GetBrandWitak2MaeipItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.socid as shopid,m.code as orderno, d.id, d.sellcash as sellprice, d.itemno,d.sellcash as realsellprice,"
		sqlStr = sqlStr + " d.imakerid as makerid,"
		sqlStr = sqlStr + " d.imakerid as jungsanid, d.iitemgubun,d.itemid as itemid,d.itemoption, d.iitemname as itemname,"
		sqlStr = sqlStr + " d.iitemoptionname as itemoptionname, d.mwgubun, oi.mwdiv as onlinemwdiv,"
		sqlStr = sqlStr + " m.executedt as shopregdate, d.buycash as suplyprice"
		sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.iitemgubun='10' and d.itemid=oi.itemid"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and d.imakerid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " and m.ipchulflag='S'"
		sqlStr = sqlStr + " and m.socid='" + FRectShopId + "'"
		sqlStr = sqlStr + " order by d.id "


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).FIdx			 = rsget("id")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("iitemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")*-1
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")

					FItemList(i).FOnlineMwDiv   = rsget("onlinemwdiv")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetBrandSellItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid,m.orderno, d.idx, d.sellprice, d.itemno,d.realsellprice, d.makerid,"
		sqlStr = sqlStr + " d.jungsanid, d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname, "
		sqlStr = sqlStr + " m.shopregdate, d.suplyprice, IsNULL(i.shopitemprice,0) shopitemprice, oi.mwdiv as onlinemwdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " on d.itemgubun=i.itemgubun and d.itemid=i.shopitemid and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=oi.itemid"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"

		'sqlStr = sqlStr + " and Year(m.shopregdate)='" + FRectJungsanYYYY + "'"
		'sqlStr = sqlStr + " and Month(m.shopregdate)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopId + "'"
		sqlStr = sqlStr + " and d.makerid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " order by d.idx"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).FIdx			 = rsget("idx")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).FJungsanId        = rsget("jungsanid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")

					FItemList(i).Fcurrentitemprice = rsget("shopitemprice")
					FItemList(i).FOnlineMwDiv = rsget("onlinemwdiv")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetBrandSellSumList()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * d.realsellprice) as subtotal, count(m.idx) as cnt, "
		sqlStr = sqlStr + " d.makerid "
		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s on s.shopid='" + FRectShopid + "' and d.makerid=s.makerid"
		end if

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectOffgubun="OFF" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop0'" + vbcrlf
		elseif FRectOffgubun="FRN" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop8'" + vbcrlf
		elseif FRectOffgubun="CAF" then
			sqlStr = sqlStr + " and Left(m.shopid,4)='cafe'" + vbcrlf
		else
			sqlStr = sqlStr + " and Left(m.shopid,10)='streetshop'" + vbcrlf
		end if

		if FRectOnlyShop<>"" then
			sqlStr = sqlStr + " and Left(m.shopid,4)<>'cafe'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by d.makerid "
		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if
		sqlStr = sqlStr + " order by subtotal desc, cnt desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FMakerid  = rsget("makerid")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("subtotal")

					if FRectShopid<>"" then
						FItemList(i).FChargeDiv = rsget("chargediv")
					end if

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetNotMatchSellChargeIDList()
		dim i,sqlStr
		sqlStr = " select top 500 d.makerid, d.idx, d.itemname, d.itemoption, d.itemno,"
		sqlStr = sqlStr + " d.realsellprice,m.shopid, i.chargeid, d.jungsanid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
		sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		sqlStr = sqlStr + " and d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " and d.jungsanid<>i.chargeid"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).Fidx  = rsget("idx")
					FItemList(i).FMakerid  = rsget("makerid")
					FItemList(i).FItemName = db2html(rsget("itemname"))
					FItemList(i).FCount = rsget("itemno")
					FItemList(i).FSum   = rsget("realsellprice")
					FItemList(i).FShopid= rsget("shopid")

					FItemList(i).FjungsanID   = rsget("chargeid")
					FItemList(i).FSelljungsanID = rsget("jungsanid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetBrandShopSellSumList()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * d.realsellprice) as subtotal, count(m.idx) as cnt, "
		sqlStr = sqlStr + " m.shopid, c.chargeuser as isbrandshop"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_chargeuser c"

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and d.jungsanid=c.chargeuser"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		'sqlStr = sqlStr + " and m.jungsanid='" + FRectJungsanId + "'"

		sqlStr = sqlStr + " group by shopid, c.chargeuser"
		sqlStr = sqlStr + " order by subtotal desc, cnt desc"

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FMakerid  = rsget("isbrandshop")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("subtotal")
					FItemList(i).FShopid= rsget("shopid")
					FItemList(i).FIsBrandShop= rsget("isbrandshop")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetDaylySellItemList()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * d.realsellprice) as subtotal, sum(d.itemno) as itemno, "
		sqlStr = sqlStr + " d.sellprice, d.realsellprice, d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		end if

		sqlStr = sqlStr + " where m.idx=d.masteridx"

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectOffgubun="OFF" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop0'" + vbcrlf
		elseif FRectOffgubun="FRN" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop8'" + vbcrlf
		elseif FRectOffgubun="CAF" then
			sqlStr = sqlStr + " and Left(m.shopid,4)='cafe'" + vbcrlf
		else
			sqlStr = sqlStr + " and Left(m.shopid,10)='streetshop'" + vbcrlf
		end if

		if FRectTerms<>"" then
			sqlStr = sqlStr + " and convert(varchar(10),m.shopregdate,20)='" + FRectTerms + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " group by d.sellprice, d.realsellprice, d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"
		sqlStr = sqlStr + " order by subtotal desc, itemno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).Fsubtotal       = rsget("subtotal")
					FItemList(i).FMakerID		 = rsget("makerid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylySellItemList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * d.realsellprice) as subtotal, sum(d.itemno) as itemno, "
		sqlStr = sqlStr + " d.sellprice, d.realsellprice, d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			sqlStr = sqlStr + " and convert(varchar(10),dateadd(hh,-5,m.shopregdate),20)='" + FRectTerms + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " group by d.sellprice, d.realsellprice, d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"
		sqlStr = sqlStr + " order by subtotal desc, itemno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).Fsubtotal       = rsget("subtotal")
					FItemList(i).FMakerID		 = rsget("makerid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylySellItemListByShopByItem()
		dim i,sqlStr

		sqlStr = " select top 1000 convert(varchar(10),shopregdate,21) as yyyymmdd, m.shopid, d.itemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + "         d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid, "
		sqlStr = sqlStr + "         sum(d.itemno) as itemno "
	        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d "
	        sqlStr = sqlStr + " where 1 = 1 "
	        sqlStr = sqlStr + " and m.idx=d.masteridx "
	        sqlStr = sqlStr + " and m.cancelyn='N' "
	        sqlStr = sqlStr + " and d.cancelyn='N' "

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemId)
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if

	        sqlStr = sqlStr + " group by convert(varchar(10),shopregdate,21), m.shopid, d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid "
	        sqlStr = sqlStr + " order by convert(varchar(10),shopregdate,21), m.shopid, d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).Fshopregdate   = rsget("yyyymmdd")
					FItemList(i).Fshopid        = rsget("shopid")
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).FMakerID	    = rsget("makerid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylySellJumunList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select m.orderno,m.totalsum,m.realsum, m.jumunmethod,m.shopregdate,m.pointuserno,"
		sqlStr = sqlStr + " d.itemname,d.itemoptionname,d.sellprice,d.realsellprice,d.itemno, d.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			sqlStr = sqlStr + " and convert(varchar(10),dateadd(hh,-5,m.shopregdate),20)='" + FRectTerms + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " order by d.idx"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellMasterDetailItem
					FItemList(i).ForderNo         = rsget("orderno")
					FItemList(i).Ftotalsum        = rsget("totalsum")
					FItemList(i).Frealsum         = rsget("realsum")
					FItemList(i).Fshopregdate	  = rsget("shopregdate")
					FItemList(i).Fjumunmethod        = rsget("jumunmethod")
					FItemList(i).Fitemname        = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice       = rsget("sellprice")
					FItemList(i).Frealsellprice   = rsget("realsellprice")
					FItemList(i).Fitemno          = rsget("itemno")
					FItemList(i).FMakerID		  = rsget("makerid")
					FItemList(i).Fpointuserno		  = rsget("pointuserno")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylySumListByJungsanID()
		dim i,sqlStr
		sqlStr = " select top 100 sum(d.itemno * d.realsellprice) as sellsum, count(m.idx) as cnt, "
		sqlStr = sqlStr + " convert(varchar(10),m.shopregdate,20) as selldate, m.shopid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"

		sqlStr = sqlStr + " and d.makerid='" + FRectJungsanID + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by convert(varchar(10),m.shopregdate,20), m.shopid"
		sqlStr = sqlStr + " order by m.shopid, m.selldate desc"

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetDaylySumList()
		dim i,sqlStr
		sqlStr = " select top 100 count(m.idx) as cnt, sum(m.realsum) as sellsum, sum(IsNull(spendmile,0)) as spendmilesum, "
		sqlStr = sqlStr + " convert(varchar(10),m.shopregdate,20) as selldate, m.shopid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " where m.idx<>0"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyShop<>"" then
			sqlStr = sqlStr + " and Left(m.shopid,4)<>'cafe'"
		end if

		sqlStr = sqlStr + " and m.cancelyn='N'"

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by convert(varchar(10),m.shopregdate,20), m.shopid"
		sqlStr = sqlStr + " order by m.shopid, m.selldate desc"

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")
					FItemList(i).FSpendMile = rsget("spendmilesum")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetReportByDanga()
		dim i,sqlStr
		sqlStr = " select (case when m.realsum < 10000 then '0~10,000'"
		sqlStr = sqlStr + " when m.realsum >= 10000 and m.realsum < 20000 then '10,000~20,000'"
		sqlStr = sqlStr + " when m.realsum >= 20000 and m.realsum < 30000 then '20,000~30,000'"
		sqlStr = sqlStr + " when m.realsum >= 30000 and m.realsum < 40000 then '30,000~40,000'"
		sqlStr = sqlStr + " when m.realsum >= 40000 and m.realsum < 50000 then '40,000~50,000'"
		sqlStr = sqlStr + " when m.realsum >= 50000 and m.realsum < 60000 then '50,000~60,000'"
		sqlStr = sqlStr + " when m.realsum >= 60000 and m.realsum < 70000 then '60,000~70,000'"
		sqlStr = sqlStr + " when m.realsum >= 70000 and m.realsum < 80000 then '70,000~80,000'"
		sqlStr = sqlStr + " when m.realsum >= 80000 and m.realsum < 90000 then '80,000~90,000'"
		sqlStr = sqlStr + " when m.realsum >= 90000 and m.realsum < 100000 then '90,000~100,000'"
		sqlStr = sqlStr + " else 'z10,0000~' end) as gubun, "
		sqlStr = sqlStr + " count(m.idx) as cnt, sum(m.realsum) as sellsum"
		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		sqlStr = sqlStr + " where m.idx<>0"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if
		sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " group by (case when m.realsum < 10000 then '0~10,000'"
		sqlStr = sqlStr + " when m.realsum >= 10000 and m.realsum < 20000 then '10,000~20,000'"
		sqlStr = sqlStr + " when m.realsum >= 20000 and m.realsum < 30000 then '20,000~30,000'"
		sqlStr = sqlStr + " when m.realsum >= 30000 and m.realsum < 40000 then '30,000~40,000'"
		sqlStr = sqlStr + " when m.realsum >= 40000 and m.realsum < 50000 then '40,000~50,000'"
		sqlStr = sqlStr + " when m.realsum >= 50000 and m.realsum < 60000 then '50,000~60,000'"
		sqlStr = sqlStr + " when m.realsum >= 60000 and m.realsum < 70000 then '60,000~70,000'"
		sqlStr = sqlStr + " when m.realsum >= 70000 and m.realsum < 80000 then '70,000~80,000'"
		sqlStr = sqlStr + " when m.realsum >= 80000 and m.realsum < 90000 then '80,000~90,000'"
		sqlStr = sqlStr + " when m.realsum >= 90000 and m.realsum < 100000 then '90,000~100,000'"
		sqlStr = sqlStr + " else 'z10,0000~' end)"
		sqlStr = sqlStr + " order by gubun"

		rsget.Open sqlStr,dbget,1
		maxt =0
		maxc =0
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FTerm  = rsget("gubun")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					maxc = maxc + FItemList(i).FCount
					maxt = maxt + FItemList(i).FSum
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetDaylySumList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select top 200 count(m.idx) as cnt, sum(m.realsum) as sellsum, "
		sqlStr = sqlStr + " sum(IsNull(spendmile,0)) as spendmilesum, sum(IsNull(gainmile,0)) as gainmilesum, "
		sqlStr = sqlStr + " convert(varchar(10),dateadd(hh,-5,m.shopregdate),20) as selldate, m.shopid"
		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if

		sqlStr = sqlStr + " where m.idx<>0"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyShop<>"" then
			sqlStr = sqlStr + " and Left(m.shopid,4)<>'cafe'"
		end if

		sqlStr = sqlStr + " and m.cancelyn='N'"

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by convert(varchar(10),dateadd(hh,-5,m.shopregdate),20), m.shopid"
		sqlStr = sqlStr + " order by m.shopid, m.selldate desc"

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")
					FItemList(i).FSpendMile = rsget("spendmilesum")
					FItemList(i).FGainMile = rsget("gainmilesum")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetJumunMasterList()
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_shopjumun_master"
		sqlStr = sqlStr + " where idx<>0"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if


		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select idx,orderno,shopid,totalsum,realsum,jumundiv,jumunmethod,"
		sqlStr = sqlStr + " shopregdate,cancelyn,regdate,shopidx"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
		sqlStr = sqlStr + " where idx<>0"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if
		sqlStr = sqlStr + " order by shopregdate desc"

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
					set FItemList(i) = new COffShopSellMasterItem
					FItemList(i).Fidx        = rsget("idx")
					FItemList(i).Forderno    = rsget("orderno")
					FItemList(i).Fshopid     = rsget("shopid")
					FItemList(i).Ftotalsum   = rsget("totalsum")
					FItemList(i).Frealsum    = rsget("realsum")
					FItemList(i).Fjumundiv   = rsget("jumundiv")
					FItemList(i).Fjumunmethod= rsget("jumunmethod")
					FItemList(i).Fshopregdate= rsget("shopregdate")
					FItemList(i).Fcancelyn   = rsget("cancelyn")
					FItemList(i).Fregdate    = rsget("regdate")
					FItemList(i).Fshopidx    = rsget("shopidx")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylyUserCardList()
		dim i,sqlStr
		sqlStr = " select top 100 count(m.idx) as cnt, sum(IsNull(shoppoint,0)) as shoppoint, "
		sqlStr = sqlStr + " convert(varchar(10),m.regdate,20) as regdate, m.regshopid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_pointuser m"
		sqlStr = sqlStr + " where m.idx<>0"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.regshopid='" + FRectShopID + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.regdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.regdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by convert(varchar(10),m.regdate,20), m.regshopid"
		sqlStr = sqlStr + " order by m.regshopid, m.regdate desc"

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FTerm  = rsget("regdate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("shoppoint")
					FItemList(i).FShopid= rsget("regshopid")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub


	public Sub GetJumunMethodReport()

		Dim sql, i, ix
		maxt = -1
		maxt2 = -1
   		maxc = -1


		sql = "select convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart," + vbcrlf
		sql = sql + " sum(m.realsum) as sumtotal, count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
		sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by  convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)" + vbcrlf
		sql = sql + " order by  convert(varchar(10),m.shopregdate,20) desc"


		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(ix) = new COffShopSellByTerm
				FItemList(ix).Fselltotal = rsget("sumtotal")

				if Not IsNull(FItemList(ix).Fselltotal) then
					maxt2 = MaxVal(maxt2,FItemList(ix).Fselltotal)
				end if

				rsget.MoveNext
				ix = ix + 1
		loop
		rsget.close



		sql = "select convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart," + vbcrlf
		sql = sql + " sum(m.realsum) as sumtotal, count(m.idx) as sellcnt,jumunmethod" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
		sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by  convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate),jumunmethod" + vbcrlf
		sql = sql + " order by  convert(varchar(10),m.shopregdate,20) desc"

''response.write sql

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
			    FItemList(i).Fsitename = rsget("yyyymmdd")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).Fdpart = rsget("dpart")
				FItemList(i).Faccountdiv = rsget("jumunmethod")

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	public Sub GetJumunMethodReportMonth()

		Dim sql, i, ix
		maxt = -1
		maxt2 = -1
   		maxc = -1


		sql = "select convert(varchar(7),m.shopregdate,20) as yyyymm," + vbcrlf
		sql = sql + " sum(m.realsum) as sumtotal, count(m.idx) as sellcnt" + vbcrlf
		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		sql = sql + " where m.shopregdate>='2003-08-01'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by  convert(varchar(7),m.shopregdate,20)" + vbcrlf
		sql = sql + " order by  convert(varchar(7),m.shopregdate,20) desc"

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(ix) = new COffShopSellByTerm
				FItemList(ix).Fselltotal = rsget("sumtotal")

				if Not IsNull(FItemList(ix).Fselltotal) then
					maxt2 = MaxVal(maxt2,FItemList(ix).Fselltotal)
				end if

				rsget.MoveNext
				ix = ix + 1
		loop
		rsget.close


		sql = "select convert(varchar(7),m.shopregdate,20) as yyyymm," + vbcrlf
		sql = sql + " sum(m.realsum) as sumtotal, count(m.idx) as sellcnt,jumunmethod" + vbcrlf
		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		sql = sql + " where m.shopregdate>='2003-08-01'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by  convert(varchar(7),m.shopregdate,20),jumunmethod" + vbcrlf
		sql = sql + " order by  convert(varchar(7),m.shopregdate,20) desc, jumunmethod"

'response.write sql

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
			    FItemList(i).Fsitename = rsget("yyyymm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).Faccountdiv = rsget("jumunmethod")

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	public Sub GetWeeklySellCount()
		Dim sql, ix

		sql = "select convert(varchar(7),m.shopregdate,20) as yyyymm" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where convert(varchar(7),m.shopregdate,20) = '" + CStr(FRectStartDay) + "'" + vbcrlf
'		sql = sql + " and convert(varchar(7),m.shopregdate,20) < '" + CStr(FRectEndDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by convert(varchar(7),m.shopregdate,20)" + vbcrlf
		sql = sql + " order by convert(varchar(7),m.shopregdate,20) desc"

		rsget.Open sql,dbget,1

		FTotalCount = rsget.RecordCount

	    redim preserve FCountList(FTotalCount)

		do until rsget.eof
				set FCountList(ix) = new COffShopSellByTerm
			    FCountList(ix).FYYYYMMDDHHNNSS = rsget("yyyymm")
				rsget.MoveNext
				ix = ix + 1
		loop
		rsget.close

	end Sub

	public Sub GetWeeklySellReport()

		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.shopregdate,20) as yyyymm," + vbcrlf
		sql = sql + " datepart(w,m.shopregdate) as dpart," + vbcrlf
		if FRectPointYN = "Y" then
		sql = sql + " sum(m.totalsum) as sumtotal," + vbcrlf
		else
		sql = sql + " sum(m.realsum) as sumtotal," + vbcrlf
		end if
		sql = sql + " count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where convert(varchar(7),m.shopregdate,20) ='" + CStr(FRectStartDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)" + vbcrlf
		sql = sql + " order by convert(varchar(10),m.shopregdate,20) desc, datepart(w,m.shopregdate) asc" + vbcrlf

'response.write sql

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
			    FItemList(i).Fsitename = rsget("yyyymm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	public Sub ShopJumunListBybestseller()
		dim sqlStr
		dim i
		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " sum(d.itemno) as sm, d.sellprice, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.makerid, d.itemoptionname"
		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m, "
			sqlStr = sqlStr + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, "
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		end if

		sqlStr = sqlStr + " where m.orderno=d.orderno"
		sqlStr = sqlStr + " and m.shopregdate >='" + CStr(FRectStartDay) + "'"
		sqlStr = sqlStr + " and m.shopregdate <'" + CStr(FRectEndDay) + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if

		if FRectOffgubun="OFF" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop0'" + vbcrlf
		elseif FRectOffgubun="FRN" then
			sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop8'" + vbcrlf
		elseif FRectOffgubun="CAF" then
			sqlStr = sqlStr + " and Left(m.shopid,4)='cafe'" + vbcrlf
		else
			sqlStr = sqlStr + " and Left(m.shopid,10)='streetshop'" + vbcrlf
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"

		sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.itemname, d.makerid, d.itemoptionname"

		if FRectOrder="bysum" then
			sqlStr = sqlStr + " order by sum(d.itemno*d.sellprice) Desc"
		elseif FRectOrder="bycnt" then
			sqlStr = sqlStr + " order by sm Desc"
		else
			sqlStr = sqlStr + " order by sm Desc"
		end if

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		maxt =0
		maxc =0
		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
				FItemList(i).FItemNo       = rsget("sm")
				FItemList(i).FItemGubun		= rsget("itemgubun")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemOption       = rsget("itemoption")
				FItemList(i).FItemCost       = rsget("sellprice")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerid		= rsget("makerid")

				maxc = maxc + FItemList(i).FItemNo
				maxt = maxt + FItemList(i).FItemNo*FItemList(i).FItemCost
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public sub SearchMallSellrePort5()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select datepart(hh,m.shopregdate) as dpart," + vbcrlf
		sql = sql + " sum(m.realsum) as sumtotal, count(m.idx) as sellcnt" + vbcrlf
		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m" + vbcrlf
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		end if

		sql = sql + " where m.shopregdate>='" + CStr(FRectStartDay) + "'"
		sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " group by datepart(hh,m.shopregdate)" + vbcrlf
		sql = sql + " order by datepart(hh,m.shopregdate) asc"

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm

				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchCategorySellrePort()
	Dim sql, i

    maxt = -1
    maxc = -1


		''#################################################
		''실 총 데이타.
		''#################################################
'			sql = "select count(d.itemno) as sellcnt, sum(d.realsellprice*d.itemno) as sumtotal" + vbcrlf
'			if FRectOldData="on" then
'				sql = sql + " from  [db_shoplog].[dbo].tbl_old_shopjumun_master m," + vbcrlf
'				sql = sql + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
'			else
'				sql = sql + " from  [db_shop].[dbo].tbl_shopjumun_master m," + vbcrlf
'				sql = sql + " [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
'			end if
'			sql = sql + " where m.orderno = d.orderno" + vbcrlf
'			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
'			sql = sql + " and m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
'			sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
'			sql = sql + " and m.cancelyn='N'" + vbcrlf
'			sql = sql + " and d.cancelyn='N'"

'			rsget.Open sql,dbget,1

'			if not rsget.eof then
'				FDayTsellsum = rsget("sumtotal")
'				FDayTea = rsget("sellcnt")
'			end if

'			rsget.close

		''#################################################
		''데이타.
		''#################################################

			sql = "select count(d.itemno) as sellcnt, sum(d.realsellprice*d.itemno) as sumtotal," + vbcrlf
			sql = sql + " s.catecdl,l.code_nm" + vbcrlf
			if FRectOldData="on" then
				sql = sql + " from  [db_shoplog].[dbo].tbl_old_shopjumun_master m," + vbcrlf
				sql = sql + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
			else
				sql = sql + " from  [db_shop].[dbo].tbl_shopjumun_master m," + vbcrlf
				sql = sql + " [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
			end if
			sql = sql + " ,[db_shop].[dbo].tbl_shop_item s" + vbcrlf
			sql = sql + " left join [db_item].[dbo].tbl_item_large l " + vbcrlf
			sql = sql + " 	on s.catecdl=l.code_large" + vbcrlf
			sql = sql + " where m.orderno = d.orderno" + vbcrlf
			sql = sql + " and d.itemgubun=s.itemgubun" + vbcrlf
			sql = sql + " and d.itemid=s.shopitemid" + vbcrlf
			sql = sql + " and d.itemoption=s.itemoption" + vbcrlf
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
			sql = sql + " and m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
			sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
			sql = sql + " and m.cancelyn='N'" + vbcrlf
			sql = sql + " and d.cancelyn='N'" + vbcrlf
			sql = sql + " group by s.catecdl,l.code_nm" + vbcrlf
			sql = sql + " order by s.catecdl"

			rsget.Open sql,dbget,1

			FResultCount = rsget.RecordCount


		    redim preserve FItemList(FResultCount)


			do until rsget.eof

				set FItemList(i) = new COffShopSellByTerm
			    FItemList(i).Fsitename = rsget("code_nm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")

				if IsNULL(FItemList(i).Fsitename) then FItemList(i).Fsitename = "미지정"

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
			loop

			rsget.close

	end sub

	Private Sub Class_Initialize()
'		redim preserve FItemList(0)
'		redim preserve FCountList(0)
		redim  FItemList(0)
		redim  FCountList(0)
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