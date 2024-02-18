<%
'####################################################
' Description :  오프라인 상품 클래스
' History : 2012.08.30 한용민 생성
'####################################################

class CoffstockItem
	public FItemGubun
	public FItemID
	public FItemOption
	public FItemName
	public FItemOptionName
	public FItemNo
	public FLastDate
	public FSellcash
	public FBuycash
	public FIsUsing
	public FSellYn
	public FLimityn
	public FLimitNo
	public FLimitSold
	public FImageSmall
	public FImageList
	public FSellNo
	public Fipno
	public Fchulno
	public Foptionusing
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fdeliverytype
	public Fregdate
	public FMakerID
	public Fpojangok
	public FMwDiv
	public FOptionCnt
	public FRackCode
	public FitemRackCode
	public Fbrandname
	public Fcurrno
	public Fsell7days
	public Fjupsu7days
	public Foffchulgo7days
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public Fpreorderno
	public FChargediv
	public FIpkumdiv4
	public FIpkumdiv2
	public Fipkumdiv5
	public Frealstock
	public FLastUpdate
	public Foldstockupdate
	public Foldstockcurrno
	public Fdanjongyn
	public FOffimgMain
	public FOffimgSmall

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function

	public Function GetDivName()
		if FMwDiv="" then
			GetDivName = "판매"
		elseif FMwDiv="ST" then
			GetDivName = "입고"
		elseif FMwDiv="SO" then
			GetDivName = "출고"
		end if
	end function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetMayNo()
		GetMayNo = FItemNo + Fipno + Fchulno - FSellNo
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if (FItemOption="0000") then
			if FLimityn="Y" then
				if FLimitNo-FLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FLimitNo-FLimitSold)
				end if
			end if
		else
			if FOptLimityn="Y" then
				if FOptLimitNo-FOptLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FOptLimitNo-FOptLimitSold)
				end if
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if FIsUsing="N" then
			reStr = reStr + " 사용x"
		end if

		if FSellYn="N" then
			reStr = reStr + " 판매x"
		end if

		if FLimityn="Y" then
			reStr = reStr + " 한정" + CStr(GetLimitEa()) + "개"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetDeliveryName()
		if Fdeliverytype="1" then
			GetDeliveryName = "자체배송"
		elseif Fdeliverytype="2" then
			GetDeliveryName = "업체배송"
		elseif Fdeliverytype="3" then
			GetDeliveryName = "?"
		elseif Fdeliverytype="4" then
			GetDeliveryName = "자체무료배송"
		elseif Fdeliverytype="5" then
			GetDeliveryName = "업체무료배송"
		else
			GetDeliveryName = "미지정"
		end if

	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "텐위"
		elseif FChargeDiv="4" then
			getChargeDivName = "텐매"
		elseif FChargeDiv="6" then
			getChargeDivName = "업위"
		elseif FChargeDiv="8" then
			getChargeDivName = "업매"
		end if
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CoffstockItemlist
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FOneItem
	public frectitemid
	public frectitemoption
	public frectitemgubun

	'//admin/stock/itemcurrentstock.asp
	public Sub GetoffItemDefaultData
		dim i, sqlStr, sqlsearch

		if frectitemgubun <> "" then
			sqlsearch = sqlsearch & " and i.itemgubun = '"&frectitemgubun&"'"
		end if
		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and i.shopitemid = "&frectitemid&""
		end if

		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and i.itemoption = '"&frectitemoption&"'"
		end if

		sqlStr = "select top 1"
		sqlStr = sqlstr + " i.itemgubun, i.shopitemid, i.shopitemname, i.makerid, i.shopitemprice, i.shopsuplycash"
		sqlStr = sqlstr + " , i.isusing, i.regdate, i.centermwdiv as mwdiv"
		sqlStr = sqlstr + " ,i.offimgmain ,i.offimglist ,i.offimgsmall , i.itemoption, i.shopitemoptionname"
		sqlstr = sqlstr + " ,IsNull(sm.realstock,0) as realstock"
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv5,0) as ipkumdiv5"
		sqlstr = sqlstr + " ,IsNull(sm.offconfirmno,0) as offconfirmno"
		sqlstr = sqlstr + " ,sm.lastupdate"
		sqlStr = sqlstr + " from db_shop.dbo.tbl_shop_item i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " 	on i.itemgubun = sm.itemgubun"
		sqlstr = sqlstr + " 	and i.shopitemid=sm.itemid"
		sqlstr = sqlstr + " 	and i.itemoption=sm.itemoption"
		sqlStr = sqlstr + " where 1=1 " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FOneItem = new CoffstockItem

				FOneItem.fitemgubun        = rsget("itemgubun")
				FOneItem.FItemID        = rsget("shopitemid")
				FOneItem.FItemOption    = rsget("itemoption")
				FOneItem.FItemName      = db2html(rsget("shopitemname"))
				FOneItem.FItemOptionName = db2html(rsget("shopitemoptionname"))
				FOneItem.FIsUsing 		= rsget("isusing")
				FOneItem.FOffimgMain	= rsget("offimgmain")
				FOneItem.FImageList	= rsget("offimglist")
				FOneItem.FOffimgSmall	= rsget("offimgsmall")

				if FOneItem.FOffimgMain<>"" then FOneItem.FOffimgMain = webImgUrl + "/offimage/offmain/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + FOneItem.FOffimgMain
				if FOneItem.FImageList<>"" then FOneItem.FImageList = webImgUrl + "/offimage/offlist/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + FOneItem.FImageList
				if FOneItem.FOffimgSmall<>"" then FOneItem.FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + FOneItem.FOffimgSmall

				FOneItem.Fregdate		= rsget("regdate")
				FOneItem.FMakerID		= rsget("makerid")
				FOneItem.FMwDiv     	= rsget("mwdiv")
				FOneItem.FSellcash     	= rsget("shopitemprice")
				FOneItem.FBuycash = rsget("shopsuplycash")
				FOneItem.Frealstock		 = rsget("realstock")
				FOneItem.Fipkumdiv5		 = rsget("ipkumdiv5")
				FOneItem.Foffconfirmno	 = rsget("offconfirmno")
				FOneItem.FLastUpdate	 = rsget("lastupdate")
		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
